Hi,

This is an updated version of the EnumerationX program, posted 6/26/2000 to FreeVBCode.com.
In addition to enumerating all windows, the program now includes a number of other features, which you can access by right-clicking the entry for the window in the left-hand side of the list view control.

I'm just right now going to explain the new feature which is spying (enumerating) menus of other applications. When you enumerate the parent windows you can right click any handle and if it has any menus, the 'Spy Menus' option will be enabled, otherwise it will be disabled.

WARNING: If the parent having menus is MDI And its child windows are MAXIMIZED Then spying the menus wil CRASH EnumeratioonX and may also crash VB too if running inthe VB IDE, but if the child windows are not maximized then the menus will be spied accurately.

When the spy window is active and the 'bring window to top' checkbox is checked then double clicking any menu item will bring the selected window to the top and show the actions of clicking the menu, however if the checkbox isn't checked then the EnumerationX will remain on top but the action of clicking will take place. It does not spy the menus of floating type toolbars like that of Word 97/2000 and IE, but it can spy the menus of Adobe Acrobat, WordPad, Calculator, Notepad, Winzip etc, the reason is that I don't seem to be getting the handle of the menus of IE or Word. Whenever I can manage to do that I'll post an upgrade.

For questions and comments please feel free to contact me at <joehacker@yahoo.com> .  I hope this program helps you understand things you didnt know.
Regards.
 - Abubakar
