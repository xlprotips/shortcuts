Introduction
============

Most Excel wizards learn every shortcut key they can get their hands on. But to
become a real Excel pro, you sometimes need even more shortcuts. This addin
defines a slew of shortcuts intended to give you more power than what Excel
offers by default. The addin started based on some of the shortcuts defined by
[F1F9][1] in their [31 day course][2], however improvements have been made to
make some of the shortcuts more general and extensible. This addin also defines
some new shortcuts as well.

To get an overview of the many of the shortcuts, you may want to take the
course to get a feel for why the various conventions used by this addin are
useful. I could explain here, but they do a pretty thorough job so I don't need
to reinvent the wheel :) Read below to find out about some of the shortcuts
defined in the addin.


Shortcuts
=========

The default shortcuts defined in this addin are shown below. Note that you can
change the shortcut to whatever you want (see 'Hacking' section below).

Navigation
----------

- `Alt-<Right>` go to cell referenced in formula of current cell (like
  following a hyperlink "forward").
- `Alt-<Left>` go back to the last cell you were on (before clicking
  `Alt-<Right>`).

You can use these shortcuts multiple times (e.g., follow several formula links
forward, then go back several cells to where you started).

Number Formatting
-----------------

- `Ctrl-3` formats the currently selected cell(s) as a number. Using this
  shortcut multiple times toggles through various formats (comma, number/no
  comma, zero {i.e., instead of "-" show "0"}, accounting {dollar sign aligned
  left}, comma$ {dollar sign next to number}). When you reach the end you can
  continue hitting the shortcut key to cycle through each style again. If you
  want to add your own styles, see the 'Hacking' section below.

- `Ctrl-Shift-h` formats the currently selected cell(s) in thousands (e.g.,
  1000 looks like a 1).

- `Ctrl-Shift-.` formats the currently selected cell(s) as a "factor" (e.g., 1
  = 1.0000).

- `Ctrl-Shift-5` formats the currently selected cell(s) as a percent. Note that
  it is different than the standard Excel percent format so that it aligns
  properly with all the other number formats herein.

- `Ctrl-Shift-/` formats the currently selected cell(s) as a date. Using this
  shortcut multiple times toggles through various formats (DD MMM YY, DD MMM
  YYYY, MM-DD-YYYY, MMM-YYYY). Like number formatting, you can define your own
  formats by reading the 'Hacking' section below.

- `Ctrl-,` increases the number of decimal places of the currently selected
  cell(s). Can (and should) be used in conjunction with any of the number
  formats above (excluding date formats of course).

- `Ctrl-Shift-,` decreases the number of decimal places of the currently
  selected cell(s).

Formatting
----------

- `Ctrl-Shift-m` formats the currently selected cell(s) with blue font
  representing an "import" to the current sheet.

- `Ctrl-Shift-x` formats the currently selected cell(s) with red font
  representing an "export" from the current sheet.

- `Ctrl-Shift-b` resets the font of the currently selected cell(s) to their
  default color (normally black).

- `Ctrl-Shift-w` colors the currently selected cell(s) with a black background
  and white font. Good for headers, etc.

- `Ctrl-Shift-i` colors the currently selected cell(s) with a light yellow
  background representing an "input" cell.

- `Ctrl-Shift-g` colors the currently selected cell(s) with a gray background.

- `Ctrl-Shift-y` colors the currently selected cell(s) with a yellow background
  (brighter than and distinctly different from the background used by
  `Ctrl-Shift-i`).

- `Ctrl-Shift-c` clears the background of the currently selected cell(s).

Miscellaneous
-------------

- `Ctrl-Shift-a` copies the current cell(s) to the right only stopping when a
  `Ctrl-rightarrow` would stop. If you hide columns as they suggest in the
  [F1F9 course][2] this becomes immensely useful. If you do it on a standard
  sheet when there is nothing to the right of the current cell, it will copy
  the current cell to the very last column in the workbook which may not be
  what you intended.

- `Ctrl-Shift-q` pastes a link to the cell(s) currently copied to the clipboard
  using row locking as suggested in the course. It also formats the current row
  in blue if you are pasting the link on a different worksheet then the copied
  cell(s), but leaves it black if it is pasted to the same sheet as the copied
  cell(s).

- `Ctrl-Shift-n` sums all the cells to the right using the same semantics as
  `Ctrl-Shift-a`. Again, if you are setting up your model as suggested in the
  course this can be a very useful shortcut. Not quite as useful when you don't
  follow the guidelines.

- `Ctrl-Shift-t` wraps the text of each of the currently selected cell(s) in
  brackets (e.g., \[Like this\]) representing a temporary placeholder that must
  be changed before the model is finalized. Using this shortcut on cell(s) that
  are already wrapped in brackets removes the brackets (i.e., this shortcut
  toggles between \[this\] and this) thus removing the "temporary status"
  assigned to it by the brackets in the first place.

- `Ctrl-F9` calculate all cells currently selected (but nothing more).
  Obviously only useful when you have "manual calculation" set in your Excel
  options.

- `Ctrl-Shift-p` resets some pivot table options (don't change column widths
  and format any values as a number using the format defined above). This was
  not ever described in the course, but is a useful shortcut for me so it got
  added here.


Excel Options
=============

The course recommends and I agree that you should change some of the default
settings in Excel to make modeling a bit easier:

- **Manual Calculation**. I tend to leave manual calculation on all the time,
  even when I am not modeling. It makes it easier to not have to think about
  whether Excel will calculate automatically or not. This is probably one of
  the more controversial recommendations if I had to guess as virtually no one
  has this set (and many people don't even know it exists), but if you can get
  others on board I think it is a useful setting.

- **Don't move down on Enter**. In Excel, if you hit `Alt t o` and click the
  "Advanced" button on the left, the first option you are faced with is what to
  do when you hit the Enter key. By default, Excel will go to the next cell,
  but it is often useful to have it do nothing (other than exit the cell you
  are editing). I.e., it should stay on the same cell you are/were editing. I
  find this to be a very useful setting, but I'm not sure why. Try it out for a
  few days and maybe you will, maybe you won't.


Hacking
=======

If you want to modify the shortcut keys used for any functionality offered in
the addin, you can do so by using a form available directly from the addin.
Simply choose the "Customize Quick Access Toolbar" and choose "More Commands".
Then under "Choose commands from:" choose "Macros". Within the list that pops
up you should see a macro called "edit_shortcuts" that you can "Add" to your
quick access toolbar. Once that has been added, simply click the button and a
form will pop up with all the shortcuts offered by this addin which can be
modified by clicking on the item you want to change, hitting the new shortcut
key you want to use in the first text box, and "Save"ing your change.

You can also add new formats to the number formatting macro and the date
formatting macro. You have to modify some code (Visual Basic), but it is not
all that difficult. Hit `Alt-F11` to open the Visual Basic Editor and make sure
the "Project Explorer" is visible (you can hit `Ctrl-r` to be sure). In the
Project Explorer you will see a file called "F1F9_macros.xla". Under the
"Modules" folder of that project, double-click the "F1F9" module and find the
"comma_style" sub therein (you can also use the dropdown menu in the upper
right hand portion of your screen to go right to this sub). You'll note that
there is an array defined in the first line of the procedure (`Dim s(0 to 6)`).
If you are going to add a new format, you will need to update this line to
account for the new style (e.g., `Dim s(0 to 7)`). Then, copy the line right
before the one that reads `Call toggle_style(s, "Comma")` and insert a new line
for your style (again right before the line that calls `toggle_style`). Name
the format something that will be meaningful to you and type the format you
want in the second string/argument. Save your work (see the save button in the
toolbar) and that's it. The same method can be followed to add new date styles
under the `toggle_date_style` procedure.


Installing
==========

Installing the addin should be as easy as downloading the xla file and opening
it up in Excel. If you want to have it opened every time you start Excel, put
it into your "startup" folder.


[1]: http://www.f1f9.com/
[2]: http://info.f1f9.com/31-day-financial-modelling-course
