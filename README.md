# Excel Vim Keys
Too cool to use Excel, too amateur to escape it\
![](./extra/typing.gif)

Excel shortcuts can be used in conjunction w/ Vim keys\
Nearly all Excel `{CNTRL}` & `{SHIFT}` shortcuts are unaffected (excl. underline & auto-fill-right)\
Excel `{ALT}` shortcuts are unaffected

Movements to leftmost & rightmost populated cells in a row is most enjoyable advantage

## Quick Setup
1. Download [vim\_keys.xlam](vim_keys.xlam)
2. Skip to step 4 below in DIY Setup

## DIY Setup
1. If Excel Developer tab isn't enabled yet, [enable it in Excel Options](https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45)
2. In a new Excel file, from Developer tab, in Visual Basic editor, create 2 modules `vim_emulation` & `vim_shortcuts` and paste in corresponding contents of sub procedures from [vim\_emulation.bas](vim_emulation.bas) & [vim\_shortcuts.bas](vim_shortcuts.bas)
3. Save-As the Excel file as `Excel Add-in (*.xlam)`
    - Default location is `~\AppData\Roaming\Microsoft\AddIns\<file-name>.xlam`
4. In an Excel file, from Developer tab, in Excel Add-ins, browse for & add the Excel Add-in
5. In Excel Options `File > Options`, in dialog box section `Quick Access Toolbar`, `Choose commands from` Macros, add `setup_shortcuts` & `teardown_shortcuts`
    - Modify names & icons to preferrence; I use `vim_mode` with checkmark icon & `stop_vim` with cancel icon

[Excel's cut is fake](https://superuser.com/questions/611854/prevent-excel-from-clearing-copied-data-for-pasting-after-certain-operations-w)

# Keys
### Normal Mode
|key|action|
|---|---|
|`h`|move left|
|`j`|move down|
|`k`|move up|
|`l`|move right|
|`{BKSP}`|move left|
|`{SPACE}`|move right|
|`i`|edit cell|
|`a`|edit cell|
|`A`|edit cell right of rightmost value in row|
|`I`|edit cell left of leftmost value in row|
|`o`|insert row below|
|`O`|insert row above|
|`x`|delete|
|`d`|cut|
|`D`|clear row's cell contents from selected to right|
|`r`|replace cell contents|
|`R`|replace cell contents|
|`b`|move contiguous left|
|`w`|move contiguous right|
|`e`|move contiguous right|
|`H`|move top of viewport|
|`{CNTRL}`+`u`|move page-up|
|`L`|move bottom of viewport|
|`{CNTRL}`+`d`|move page-down|
|`{SHIFT}`+`4`|move to rightmost value in row|
|`0`|move to column `A` in row|
|`_`|move to leftmost value in row|
|`^`|move to leftmost value in row|
|`v`|start `visual` mode|
|`p`|paste|
|`P`|paste values|
|`u`|undo|
|`{CNTRL}`+`r`|redo|
|`/`|search|
|`n`|next search result|
|`N`|previous search result|

### Visual Mode
|key|action|
|---|---|
|`h`|move left|
|`j`|move down|
|`k`|move up|
|`l`|move right|
|`{BKSP}`|move left|
|`{SPACE}`|move right|
|`b`|move contiguous left|
|`w`|move contiguous right|
|`e`|move contiguous right|
|`$`|move to rightmost value in row|
|`0`|move to column `A` in row|
|`_`|move to leftmost value in row|
|`^`|move to leftmost value in row|
|`x`|delete|
|`d`|cut|
|`y`|copy|
|`p`|paste|
|`P`|paste values|
|`{CNTRL}`+`u`|move page-up|
|`{CNTRL}`+`d`|move page-down|
|`v`|exit visual mode to normal mode|
|`{ESC}`|exit visual mode to normal mode|

---

## Reference Projects
- [ExcelLikeVim](https://github.com/kjnh10/ExcelLikeVim)
- [xlpro.tips](https://xlpro.tips/posts/excel-and-vim/)
- [vim\_ahk](https://github.com/rcmdnk/vim_ahk)
- [vibre\_office](https://github.com/seanyeh/vibreoffice)

## VBA API Docs
- [Application.OnKey](https://learn.microsoft.com/en-us/office/vba/api/excel.application.onkey)
- [Application.SendKeys](https://learn.microsoft.com/en-us/office/vba/api/excel.application.sendkeys)

