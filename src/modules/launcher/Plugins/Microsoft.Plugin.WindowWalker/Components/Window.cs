// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

// Code forked from Betsegaw Tadele's https://github.com/betsegaw/windowwalker/
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Wox.Plugin.Common.VirtualDesktop.Helper;
using Wox.Plugin.Common.Win32;
using Wox.Plugin.Logger;

namespace Microsoft.Plugin.WindowWalker.Components
{
    /// <summary>
    /// Represents a specific open window
    /// </summary>
    internal class Window
    {
        /// <summary>
        /// Specifies the maximum allowed length, in characters, for a window class name,
        /// including trailing null.
        /// </summary>
        private const int MaxClassNameLength = 256;

        /// <summary>
        /// The handle to the window
        /// </summary>
        private readonly IntPtr hwnd;

        /// <summary>
        /// Caches the process data for enumerated windows. Access must be guarded via a lock.
        /// </summary>
        private static readonly Dictionary<IntPtr, WindowProcess> _handlesToProcessCache = [];

        /// <summary>
        /// An instance of <see cref="WindowProcess"/> that contains the process information for the window
        /// </summary>
        private readonly WindowProcess processInfo;

        /// <summary>
        /// An instance of <see cref="VDesktop"/> that contains the desktop information for the window
        /// </summary>
        private readonly VDesktop desktopInfo;

        /// <summary>
        /// Limits concurrent expensive UWP window name fixup tasks (child enumeration
        /// and process queries). Prevents queuing too many concurrent fixups when typing
        /// quickly and/or when many UWP windows are open.
        /// </summary>
        private static readonly SemaphoreSlim _uwpFixupSemaphore = new(2, 2);

        /// <summary>
        /// Holds a thread-local cached instance of a <see cref="StringBuilder"/> for
        /// reuse within the current thread.
        /// </summary>
        [ThreadStatic]
        private static StringBuilder _cachedBuilder;

        /// <summary>
        /// Helper to retrieve a cached StringBuilder instance for the current thread
        /// with at least the requested capacity.
        /// </summary>
        private static StringBuilder GetCachedStringBuilder(int capacity)
        {
            _cachedBuilder ??= new StringBuilder(capacity);

            _cachedBuilder.Clear();

            if (_cachedBuilder.Capacity < capacity)
            {
                _cachedBuilder.EnsureCapacity(capacity);
            }

            return _cachedBuilder;
        }

        /// <summary>
        /// Gets the title of the window (the string displayed at the top of the window)
        /// </summary>
        internal string Title
        {
            get
            {
                int length = NativeMethods.GetWindowTextLength(hwnd);
                if (length > 0)
                {
                    var builder = GetCachedStringBuilder(length + 1);
                    length = NativeMethods.GetWindowText(hwnd, builder, builder.Capacity);
                    if (length > 0)
                    {
                        return builder.ToString();
                    }
                }

                return string.Empty;
            }
        }

        /// <summary>
        /// Gets the handle to the window
        /// </summary>
        internal IntPtr Hwnd
        {
            get { return hwnd; }
        }

        /// <summary>
        /// Gets the object of with the process information of the window
        /// </summary>
        internal WindowProcess Process
        {
            get { return processInfo; }
        }

        /// <summary>
        /// Gets the object of with the desktop information of the window
        /// </summary>
        internal VDesktop Desktop
        {
            get { return desktopInfo; }
        }

        /// <summary>
        /// Gets the name of the class for the window represented
        /// </summary>
        internal string ClassName
        {
            get
            {
                return GetWindowClassName(Hwnd);
            }
        }

        /// <summary>
        /// Gets a value indicating whether the window is visible (might return false if it is a hidden IE tab)
        /// </summary>
        internal bool Visible
        {
            get
            {
                return NativeMethods.IsWindowVisible(Hwnd);
            }
        }

        /// <summary>
        /// Gets a value indicating whether the window is cloaked (true) or not (false).
        /// (A cloaked window is not visible to the user. But the window is still composed by DWM.)
        /// </summary>
        internal bool IsCloaked
        {
            get
            {
                return GetWindowCloakState() != WindowCloakState.None;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the specified window handle identifies an existing window.
        /// </summary>
        internal bool IsWindow
        {
            get
            {
                return NativeMethods.IsWindow(Hwnd);
            }
        }

        /// <summary>
        /// Gets a value indicating whether the window is a toolwindow
        /// </summary>
        internal bool IsToolWindow
        {
            get
            {
                return (NativeMethods.GetWindowLong(Hwnd, Win32Constants.GWL_EXSTYLE) &
                    (uint)ExtendedWindowStyles.WS_EX_TOOLWINDOW) ==
                    (uint)ExtendedWindowStyles.WS_EX_TOOLWINDOW;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the window is an appwindow
        /// </summary>
        internal bool IsAppWindow
        {
            get
            {
                return (NativeMethods.GetWindowLong(Hwnd, Win32Constants.GWL_EXSTYLE) &
                    (uint)ExtendedWindowStyles.WS_EX_APPWINDOW) ==
                    (uint)ExtendedWindowStyles.WS_EX_APPWINDOW;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the window has ITaskList_Deleted property
        /// </summary>
        internal bool TaskListDeleted
        {
            get
            {
                return NativeMethods.GetProp(Hwnd, "ITaskList_Deleted") != IntPtr.Zero;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the specified windows is the owner (i.e. doesn't have an owner)
        /// </summary>
        internal bool IsOwner
        {
            get
            {
                return NativeMethods.GetWindow(Hwnd, GetWindowCmd.GW_OWNER) == IntPtr.Zero;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the window is minimized
        /// </summary>
        internal bool Minimized
        {
            get
            {
                return GetWindowSizeState() == WindowSizeState.Minimized;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Window"/> class.
        /// Initializes a new Window representation
        /// </summary>
        /// <param name="hwnd">the handle to the window we are representing</param>
        internal Window(IntPtr hwnd)
        {
            // TODO: Add verification as to whether the window handle is valid
            this.hwnd = hwnd;
            processInfo = CreateWindowProcessInstance(hwnd);
            desktopInfo = Main.VirtualDesktopHelperInstance.GetWindowDesktop(hwnd);
        }

        /// <summary>
        /// Switches desktop focus to the window
        /// </summary>
        internal void SwitchToWindow()
        {
            // The following block is necessary because
            // 1) There is a weird flashing behavior when trying
            //    to use ShowWindow for switching tabs in IE
            // 2) SetForegroundWindow fails on minimized windows
            // Using Ordinal since this is internal
            if (processInfo.Name.ToUpperInvariant().Equals("IEXPLORE.EXE", StringComparison.Ordinal) || !Minimized)
            {
                NativeMethods.SetForegroundWindow(Hwnd);
            }
            else
            {
                if (!NativeMethods.ShowWindow(Hwnd, ShowWindowCommand.Restore))
                {
                    // ShowWindow doesn't work if the process is running elevated: fall back to SendMessage
                    _ = NativeMethods.SendMessage(Hwnd, Win32Constants.WM_SYSCOMMAND, Win32Constants.SC_RESTORE);
                }
            }

            NativeMethods.FlashWindow(Hwnd, true);
        }

        /// <summary>
        /// Helper function to close the window
        /// </summary>
        internal void CloseThisWindowHelper()
        {
            _ = NativeMethods.SendMessageTimeout(Hwnd, Win32Constants.WM_SYSCOMMAND, Win32Constants.SC_CLOSE, 0, 0x0000, 5000, out _);
        }

        /// <summary>
        /// Closes the window
        /// </summary>
        internal void CloseThisWindow()
        {
            Thread thread = new(new ThreadStart(CloseThisWindowHelper));
            thread.Start();
        }

        /// <summary>
        /// Converts the window name to string along with the process name
        /// </summary>
        /// <returns>The title of the window</returns>
        public override string ToString()
        {
            // Using CurrentCulture since this is user facing
            return Title + " (" + processInfo.Name.ToUpper(CultureInfo.CurrentCulture) + ")";
        }

        /// <summary>
        /// Returns what the window size is
        /// </summary>
        /// <returns>The state (minimized, maximized, etc..) of the window</returns>
        internal WindowSizeState GetWindowSizeState()
        {
            NativeMethods.GetWindowPlacement(Hwnd, out WINDOWPLACEMENT placement);

            switch (placement.ShowCmd)
            {
                case ShowWindowCommand.Normal:
                    return WindowSizeState.Normal;
                case ShowWindowCommand.Minimize:
                case ShowWindowCommand.ShowMinimized:
                    return WindowSizeState.Minimized;
                case ShowWindowCommand.Maximize: // No need for ShowMaximized here since its also of value 3
                    return WindowSizeState.Maximized;
                default:
                    // throw new Exception("Don't know how to handle window state = " + placement.ShowCmd);
                    return WindowSizeState.Unknown;
            }
        }

        /// <summary>
        /// Enum to simplify the state of the window
        /// </summary>
        internal enum WindowSizeState
        {
            Normal,
            Minimized,
            Maximized,
            Unknown,
        }

        /// <summary>
        /// Returns the window cloak state from DWM
        /// (A cloaked window is not visible to the user. But the window is still composed by DWM.)
        /// </summary>
        /// <returns>The state (none, app, ...) of the window</returns>
        internal WindowCloakState GetWindowCloakState()
        {
            _ = NativeMethods.DwmGetWindowAttribute(Hwnd, (int)DwmWindowAttributes.Cloaked, out int isCloakedState, sizeof(uint));

            switch (isCloakedState)
            {
                case (int)DwmWindowCloakStates.None:
                    return WindowCloakState.None;
                case (int)DwmWindowCloakStates.CloakedApp:
                    return WindowCloakState.App;
                case (int)DwmWindowCloakStates.CloakedShell:
                    return Main.VirtualDesktopHelperInstance.IsWindowCloakedByVirtualDesktopManager(hwnd, Desktop.Id) ? WindowCloakState.OtherDesktop : WindowCloakState.Shell;
                case (int)DwmWindowCloakStates.CloakedInherited:
                    return WindowCloakState.Inherited;
                default:
                    return WindowCloakState.Unknown;
            }
        }

        /// <summary>
        /// Enum to simplify the cloak state of the window
        /// </summary>
        internal enum WindowCloakState
        {
            None,
            App,
            Shell,
            Inherited,
            OtherDesktop,
            Unknown,
        }

        /// <summary>
        /// Returns the class name of a window.
        /// </summary>
        /// <param name="hwnd">Handle to the window.</param>
        /// <returns>Class name</returns>
        private static string GetWindowClassName(IntPtr hwnd)
        {
            var builder = GetCachedStringBuilder(MaxClassNameLength);
            return NativeMethods.GetClassName(hwnd, builder, builder.Capacity) == 0 ? string.Empty : builder.ToString();
        }

        /// <summary>
        /// Attempts to identify and update the child process information for a UWP
        /// application's main window.
        /// </summary>
        /// <param name="windowHandle">A handle to the main window of the process to
        /// inspect. Must be a valid window handle.</param>
        /// <param name="originalProcessInfo">The original process information associated
        /// with the specified window. Used to verify and update process details if a UWP
        /// child window is found.</param>
        private static void RunUwpFixup(IntPtr windowHandle, WindowProcess originalProcessInfo)
        {
            var acquiredSemaphore = false;
            var found = false;
            var childProcessId = 0u;
            var childThreadId = 0u;
            var childProcessName = string.Empty;

            try
            {
                _uwpFixupSemaphore.Wait();
                acquiredSemaphore = true;

                EnumWindowsProc callbackptr = new((hwnd, lParam) =>
                {
                    // Search for the child window that belongs to the UWP app process,
                    // ignoring "ApplicationFrame" windows.
                    if (GetWindowClassName(hwnd).StartsWith("Windows.UI.Core.", StringComparison.OrdinalIgnoreCase))
                    {
                        // Retrieve the information for the child window's process. The
                        // information is committed back to the cache after enumeration.
                        // We use locals here to avoid a lock on the cache during
                        // enumeration.
                        childProcessId = WindowProcess.GetProcessIDFromWindowHandle(hwnd);
                        childThreadId = WindowProcess.GetThreadIDFromWindowHandle(hwnd);
                        childProcessName = WindowProcess.GetProcessNameFromProcessID(childProcessId);

                        found = true;
                        return false;   // stop enumeration
                    }

                    return true;
                });

                _ = NativeMethods.EnumChildWindows(windowHandle, callbackptr, 0);
            }
            finally
            {
                if (acquiredSemaphore)
                {
                    _uwpFixupSemaphore.Release();
                }

                lock (_handlesToProcessCache)
                {
                    // Cache entries may have been evicted, so verify that the original
                    // process info is still associated with the window handle.
                    if (_handlesToProcessCache.TryGetValue(windowHandle, out WindowProcess processInfo)
                        && ReferenceEquals(processInfo, originalProcessInfo))
                    {
                        if (found)
                        {
                            processInfo.UpdateProcessInfo(childProcessId, childThreadId, childProcessName);
                        }

                        processInfo.IsUwpFixupInFlight = false;
                    }
                }
            }
        }

        /// <summary>
        /// Gets an instance of <see cref="WindowProcess"/> form process cache or creates a new one. A new one will be added to the cache.
        /// </summary>
        /// <param name="hWindow">The handle to the window</param>
        /// <returns>A new Instance of type <see cref="WindowProcess"/></returns>
        private static WindowProcess CreateWindowProcessInstance(IntPtr hWindow)
        {
            WindowProcess processInfo;
            var shouldRunFixup = false;

            lock (_handlesToProcessCache)
            {
                if (_handlesToProcessCache.Count > 7000)
                {
                    Debug.Print("Clearing Process Cache because its size is " + _handlesToProcessCache.Count);
                    _handlesToProcessCache.Clear();
                }

                // Add window's process to cache if missing
                if (!_handlesToProcessCache.ContainsKey(hWindow))
                {
                    // Get process ID and name
                    var processId = WindowProcess.GetProcessIDFromWindowHandle(hWindow);
                    var threadId = WindowProcess.GetThreadIDFromWindowHandle(hWindow);
                    var processName = WindowProcess.GetProcessNameFromProcessID(processId);

                    if (processName.Length != 0)
                    {
                        _handlesToProcessCache.Add(hWindow, new WindowProcess(processId, threadId, processName));
                    }
                    else
                    {
                        // For the dwm process we cannot receive the name. This is no problem because the window isn't part of result list.
                        Log.Debug($"Invalid process {processId} ({processName}) for window handle {hWindow}.", typeof(Window));
                        _handlesToProcessCache.Add(hWindow, new WindowProcess(0, 0, string.Empty));
                    }
                }

                processInfo = _handlesToProcessCache[hWindow];

                // Marks the UWP fixup as in-flight (and throttles repeat attempts) if a
                // fixup is needed.
                shouldRunFixup = processInfo.TryBeginUwpFixup();
            }

            if (shouldRunFixup)
            {
                // Correct the process data if the window belongs to a UWP app hosted by
                // 'ApplicationFrameHost.exe'. This is a best-effort fixup that works
                // for non-minimized UWP windows. (Minimized UWP windows do not have the
                // required child window.)
                _ = Task.Run(() => RunUwpFixup(hWindow, processInfo));
            }

            return processInfo;
        }
    }
}
