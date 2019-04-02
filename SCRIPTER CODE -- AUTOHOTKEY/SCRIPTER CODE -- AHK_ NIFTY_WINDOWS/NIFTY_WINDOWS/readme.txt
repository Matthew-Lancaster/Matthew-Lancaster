
  NIFTYWINDOWS VERSION 0.9.3.1

Company:  Enovatic-Solutions (IT Service Provider)
Author:   Oliver Pfeiffer, Bremen (GERMANY)
Homepage: http://www.enovatic.org/products/niftywindows/
Email:    niftywindows@enovatic.org


    INTRODUCTION

Welcome to NiftyWindows! This free tool provides many helpful features
for an easier handling of your Windows (c) operating system. It does not
need any complex installation nor does it read or save any information
on your system other than the application dependent files itself. The
release package includes only the main executable and some additional
text files. Redistribution of NiftyWindows and linking to the website
are welcome and strongly encouraged. In other words, feel free to share
this tool with the public, provided that the enclosed license is met.

Enovatic-Solutions is an IT Service Provider. We create custom solutions 
for professional users in all fields of Information Technology. If you 
have any demands or specific requirements, please feel free to let 
us know.

The various features of NiftyWindows are based on innovative ideas
reported to us by countless users. These ideas were adapted and
implemented from the scratch to provide maximum usability and
robustness. We really appreciate to receive as much user feedback on
common usability or innovative ideas for new features/improvements as
possible. Our motivation has been driven by user comments so do not
hesitate to contact us. Special thanks go to Chris Mallett for 
implementing and sharing the free scripting language AutoHotkey that 
provides the framework behind NiftyWindows.

Note: This software is free and will remain free forever. Unlike most
other tools the tool itself is the target, and not just a method to
reach another target (like money). NiftyWindows is not designed to be
commercial software. We designed it to meet certain needs of our own and
believe that others may find it useful as well. This program is free
software; you can redistribute it and/or modify it under the terms of
the GNU General Public License (GPL) as published by the Free Software
Foundation; either version 2 of the License, or (at your option) any
later version.


    NOTES

To use the middle, fourth and fifth mouse button you have to reset the
mouse button features in your mouse driver settings to the defaults,
otherwise NiftyWindows cannot catch the mouse events. Use AutoScroll for
the middle and Previous/Next for the additional buttons.

NiftyWindows provides several implicit features that do not need a
certain trigger action. These indispensable features can be enabled or
disabled in the tray icon menu. Therefore take a look in the menu of a
running instance to get to know these self-explanatory features.
Unfortunately most of them are not yet described in this document.

If you want to keep the original behaviour of any mouse button you can
disable the mouse hooks for every button separately in the responsible
tray icon menu. This makes sense if you need some special handling
provided by some applications like Internet Explorer or Mozilla/Firefox
(e.g. previous/next on fourth and fifth mouse button or close-tab on
middle). If you disable a mouse hook the responsible button will not
trigger the assigned NiftyWindows features. Due to this disadvantage we
recommend to configure an exceptional handling for certain applications
in your mouse driver (if possible) instead of disabling the buttons
completely. The effect is the same but you can still use the full power
of NiftyWindows with all other windows. The disabled hooks are stored in
the applications ini file to make them persistent (same directory as the
executable).

There are some enhanced features that are only supported on Windows
2000/XP or later. (e.g. window and pixel transparency as described below).


    INSTALLATION

If you choose the installer certainly everything is already completed
and the installer did the job for you. Otherwise you just have to unzip
the archive in any directory of your choice (e.g. C:\Program Files\).
After successful extraction start the main executable file. The tool is
activated and provides the nifty features immediately after execution.

Note: If you want to use NiftyWindows permanently you should add a link
to your startup group in the start menu to activate NiftyWindows
automatically after every successful system start.


    DEINSTALLATION

To remove this tool exit it first (press /WIN+X/) and then just remove
the provided files. There are no registry entries or any other files
outside of the NiftyWindows base directory.


    FEATURES

The main idea of this tool was triggered by my nice colleague Mark. He
is a confirmed Linux user since his very early teens. So he'll always
start to rail against the tricky window handling implemented in Redmond,
WA if he is pained to use some widespread software products.

The main intention of NiftyWindows is the provision of an easy user
control of all basic window interactions like dragging, resizing,
maximizing, minimizing, closing and many others. The idea is to perform
the central interactions with a single hand without any need to press
some tricky key combinations like other tools or window managers enforce.
One of the mostly unused mouse features is the press and drag of the
/RIGHT_BUTTON/. Therefore this mouse action is used for the main
features as explained below. The superior principle is to keep all
features efficient but simple.

Note: The basic window interactions mentioned above cannot be applied to
all windows because there are some special treatments (e.g. context
menus or maximized windows). If you really like to enforce the power of
NiftyWindows to these windows also, you can press (and hold) /CTRL/
*before* you start using a certain feature to manipulate these windows
as well. With this forced mode you can drag and resize the given (see
below) "special windows" (makes sense) as well as adjust the bounds of
context menus and other (naturally) unchangeable GUI components (be
careful).

During the development process many innovative ideas came up (and were
implemented certainly). So we had to choose further keys to provide some
simple trigger combinations to you. We decided to use the four most
left-bottom keys (/WIN/, /CTRL/, /SHIFT/, /ALT/) on your keyboard as
trigger combination only, to keep it simple.

The following list contains many features and modifiers. A modifier is
used in combination with a feature action to provide some additional
control. If there are more than one modifier specified for a single
feature, certainly all of the given modifiers are freely combinable to
provide the maximum flexibility in all situations.

This tool provides the following features to significantly improve the
user-interface-flow:

    * /CTRL+ALT+BACKSPACE/
    * /CTRL+WIN+BACKSPACE/
      This powerful hotkey removes *all* visual effects (like on exit)
      that have been made before by NiftyWindows. You can use this
      action as a fall-back solution to quickly revert any
      always-on-top, rolled windows and transparency features you've set
      before.

      Note: This hotkey is particularly important during your test stage
      to revert some nifty visual effects after you completed a single
      feature test and before you're going to test the next. So keep it
      in mind!

    * /RIGHT_BUTTON+DRAG/
      This is the most powerful feature of NiftyWindows. The area of
      every window is tiled in a virtual 9-cell grid with three columns
      and rows. The center cell is the largest one and you can grab and
      move a window around by clicking and holding it with the right
      mouse button. The other eight corner cells are used to resize a
      resizable window in the same manner.

      Note: This drag and resize feature is not provided on some certain
      windows because this mouse behaviour may have a special handling
      controlled by the application itself. You can use the forced mode
      (see explanation in the introduction of 'features') to manipulate
      these windows as well. Currently the following windows (and its
      possible child windows) have been taken into account:

      Windows Explorer; Cabinet; Internet Explorer; Mozilla/Firefox;
      Opera; Xplore

          o /CTRL+RIGHT_BUTTON+DRAG/
            This modifier enables the forced mode for this action as
            described in the introduction of 'features'.

          o /SHIFT+RIGHT_BUTTON+DRAG/
            This modifier lets the window bounds snap to a virtual 10
            pixel grid to ease window alignment and to allow a complex
            window arrangement.

          o /WIN+RIGHT_BUTTON+DRAG/
            This modifier lets you resize all four edges of a window in
            a symmetrical manner around the window center.

          o /ALT+RIGHT_BUTTON+DRAG/
            This modifier locks the width and height ratio (aka aspect
            ratio) during the resize operation.

    * /RIGHT_BUTTON+LEFT_BUTTON/
      Minimizes the selected window (if minimizable) to the task bar. If
      you press the left button over the titlebar the selected window
      will be rolled up instead of being minimized. You have to apply
      this action again to roll the window back down.

      Note: If you exit NiftyWindows all rolled windows will be unrolled
      first.

          o /CTRL+RIGHT_BUTTON+LEFT_BUTTON/
            This modifier enables the forced mode for this action as
            described in the introduction of 'features'.

    * /CTRL+WIN+R/
      Unrolls all windows (like on exit) that have been rolled up before
      by NiftyWindows.

    * /RIGHT_BUTTON+MIDDLE_BUTTON/
      Closes the selected window (if closeable) as if you click the
      close button in the titlebar. If you press the middle button over
      the titlebar the selected window will be sent to the bottom of the
      window stack instead of being closed.

          o /CTRL+RIGHT_BUTTON+MIDDLE_BUTTON/
            This modifier enables the forced mode for this action as
            described in the introduction of 'features'.

    * /RIGHT_BUTTON+WHEEL/
      Provides a quick task switcher (alt-tab-menu) controlled by the
      mouse wheel.

      Note: If you like task switching by using the alt-tab-menu you
      should give Microsoft's 'Alt-Tab Replacement' (part of the
      PowerToys) a try. With this PowerToy, in addition to seeing the
      icon of the application window you are switching to, you will also
      see a preview of the page. This helps particularly when multiple
      sessions of an application are open.

      http://www.microsoft.com/windowsxp/downloads/powertoys/xppowertoys.mspx

    * /MIDDLE_BUTTON/
      Initiates a double click.

    * /FOURTH_BUTTON/
      This additional button is used to toggle the windows start menu.

    * /FIFTH_BUTTON/
      This additional button is used to toggle the maximize state of the
      active window (if maximizable).

          o /CTRL+FIFTH_BUTTON/
            This modifier enables the forced mode for this action as
            described in the introduction of 'features'.

    * /WIN+0..9/
      Opens or closes an installed CD/DVD-ROM reader/writer drive tray.
      The drives are assigned to their hotkeys by the certain drive
      number in your system. The hotkeys are used in the sequence of the
      key placement on your physical keyboard from left to right (1
      refers to the first and 0 to the tenth drive). So you are limited
      to a total number of 10 drives controlled by NiftyWindows.

      Example: If you press /WIN+1/ the first drive found (ordered by
      their drive letters) will be opened or closed (depends on the
      current state).

      Note: During the open or close operation NiftyWindows is suspended
      (and waiting) so you can control only a single drive per time. If
      you want to open or close more than one drive you have to do it
      sequentially.

    * /PAUSE/
      Toggles the muteness of an installed audio card.

    * /WIN+S/
      Starts the user defined screensaver (password protection aware).

    * /CTRL+WIN+S/
      This does the same as the feature described above with the
      addition that the display(s) will be powered down shortly (five
      seconds) after the screensaver started successfully. This feature
      is very useful if you already know that you are going to leave the
      system for a long period of time.

      Note: If you perform any user input that ends the saver in the
      first seconds certainly the display(s) will not be powered down.
      After the display(s) were powered down you can wake them up by any
      key or mouse inputs as well known.

    * /WIN+LEFT_BUTTON/
    * /WIN+^/
      Toggles the always-on-top attribute of the selected/active window.

      Note: If you exit NiftyWindows any always-on-top attributes will
      be removed first.

    * /CTRL+WIN+^/
      Removes any always-on-top attributes (like on exit) that have been
      set before by NiftyWindows.

    * /WIN+WHEEL/
      Adjusts the transparency of the active window in ten percent steps
      (opaque = 100%) which allows the contents of the windows behind it
      to shine through. If the window is completely transparent (0%) the
      window is still there and clickable. If you loose a transparent
      window it will be extremly complicated to find it again because
      it's invisible (see the first hotkey in this list for emergency
      help in such situations).

      Note: This mode is called window-transparency. This feature is
      only supported on Windows 2000/XP or later and needs a powerful
      processor and video card installed in your system. If you exit
      NiftyWindows any transparency effects will be removed first.

          o /SHIFT+WIN+WHEEL/
            This modifier decreases the percent steps to one (instead of
            ten) for a more accurate control.

    * /WIN+CTRL+LEFT_BUTTON/
      Makes all pixels of the same color the mouse cursor points at
      invisible inside the target window, which allows the contents of
      the windows behind it to show through. If the user clicks on an
      invisible pixel, the click will "fall through" to the window
      behind it. You should always be concentrated by using this
      feature. Otherwise you can loose some windows completely (see the
      first hotkey in this list for emergency help in such situations).

      Note: This mode is called pixel-transparency. This feature is only
      supported on Windows 2000/XP or later and needs a powerful
      processor and video card installed in your system. If you exit
      NiftyWindows any transparency effects will be removed first.

    * /WIN+CTRL+MIDDLE_BUTTON/
      Provides a special combination of multiple features
      (window-transparency & pixel-transparency & always-on-top). This
      action is intended to be used on small dialogs (e.g. file-copy or
      download with progress bar). By using this feature you can modify
      an important dialog to stay inside your field of vision all the
      time without affecting your concentration. You have to revert
      these changes manually by using the well known features of
      NiftyWindows (see the first hotkey in this list for emergency help
      in such situations).

      Note: This mode combines 25% window-transparency and the
      pixel-transparency (pixel color at mouse click position). These
      features are only supported on Windows 2000/XP or later and needs
      a powerful processor and video card installed in your system. If
      you exit NiftyWindows any transparency effects will be removed first.

    * /WIN+MIDDLE_BUTTON/
      Removes any transparency effects (windows- as well as
      pixel-transparency) of the selected window.

      Note: If you use pixel transparency you have to click on any
      remaining visible part of the window otherwise the action is
      catched by the window behind it as described above.

    * /CTRL+WIN+T/
      Removes any transparency effects (like on exit) that have been set
      before by NiftyWindows.

    * /ALT+WHEEL/
      Changes the size of the active window in ten percent steps which
      allows a very quick resize operation with simply two snatches.

          o /CTRL+ALT+WHEEL/
            This modifier enables the forced mode for this action as
            described in the introduction of 'features'.

          o /SHIFT+ALT+WHEEL/
            This modifier decreases the percent steps to one (instead of
            ten) for a more accurate control.

          o /WIN+ALT+WHEEL/
            This modifier lets you resize all four edges of a window in
            a symmetrical manner around the window center.

    * /ALT+NumAdd/
    * /ALT+NumSub/
      Changes the size of the active window in steps of the standard
      screen resolutions. /ALT+NumAdd/ increases to the next higher,
      /ALT+NumSub/ decreases to the previous lower standard resolution.
      This feature is very useful for testing web pages: you can easily
      set your browser window to any common screen resolution.

          o /CTRL+ALT+NumAdd/
          o /CTRL+ALT+NumSub/
            This modifier enables the forced mode for this action as
            described in the introduction of 'features'.

    * /WIN+F1..F24/
      Activates the next window in a process window group that was
      defined gradually before with the given /CTRL/ modifier. This
      feature causes the first window of the responsible group to be
      activated. Using it a second time will activate the next window in
      the series and so on. By using process window groups you can
      organize and access your process windows in semantic groups quickly.

      Note: When a window is activated immediately after another window
      was activated, the task bar buttons may start flashing on some
      systems (depending on OS and settings).

          o /CTRL+WIN+F1..F24/
            This modifier adds all windows of the active process to the
            responsible group.

          o /ALT+WIN+F1..F24/
            This modifier closes *all* windows of the responsible group.

    * /WIN+ESC/
      By pressing this hotkey you can disable (and reenable) all
      NiftyWindows mouse and keyboard features at once. This is commonly
      used when you're going to start a special application that handles
      some triggers itself or when you want to temporarily disable
      NiftyWindows.

      Note: There is an auto suspend mode that can be enabled/disabled
      in the tray icon menu. This auto mode disables NiftyWindows
      temporarily as soon as a fullscreen window (propably a 3D game)
      was activated (gain focus). After the fullscreen window was
      deactivated (lost focus) NiftyWindows is enabled again.

    * /WIN+X/
      Exits NiftyWindows.

      Note: If the tray icon is hidden this hotkey will show it again
      (if it is already visible NiftyWindows will exit as expected).

------------------------------------------------------------------------

The following features are less common because they depend on special
applications and tools which *may* be installed on your system:

    * /CTRL+SHIFT+B/
      Toggles the visibility of the Miranda buddy list (if installed).
      Currently Miranda does not provide a hotkey to activate the buddy
      list if the window is still visible. Instead the opened (but not
      activated) buddy list will be minimized. This is not expected so
      this NiftyWindows feature provides the needed service asked by so
      many people:

          o If Miranda's buddy list window is visible but not activated
            the window will be activated by pressing this hotkey.
          o If the buddy list is visible and activated the window will
            be minimized.
          o If the buddy list is minimized the window will be made
            visible and activated.
          o If Miranda is not started at all it will be started and the
            buddy list will be made visible and activated.

      Note: This feature is only applied if the Miranda executable can
      be found at: %ProgramFiles%\Miranda\Miranda32.exe

      http://www.miranda-im.org/

      As soon as this service will be provided by Miranda itself this
      hotkey will be removed from NiftyWindows.

    * /CTRL+SHIFT+U/
      Toggles the visibility of the last used Miranda message container
      (if installed). Currently Miranda does not provide a hotkey to
      activate the last used message container if there is no unread
      message waiting for your attention. So this hotkey will make a
      container visible (if it is minimized) and activate it. If there
      is no existing message container, this hotkey will do nothing.

      Note: The message container identification is done by its window
      class only. So it may be that there are further windows having the
      same class that will be triggered by this hotkey also. There is
      currently no other way to do this. But you can inform us about
      some wrongly triggered window titles to exclude them from this
      hotkey in particular. Currently the following window titles have
      been taken into account:

      Mail

      This feature is only applied if the Miranda executable can be
      found at: %ProgramFiles%\Miranda\Miranda32.exe

      http://www.miranda-im.org/

      As soon as this service will be provided by Miranda itself this
      hotkey will be removed from NiftyWindows.

