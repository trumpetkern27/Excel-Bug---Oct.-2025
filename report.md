# Spooky Ghost Tooltips in Excel: A Debugging Case Study

*Posted October 4, 2025*

Recently, I stumbled on an odd Excel behavior at work: ghost tooltips that would appear when switching sheets. Normally tooltips in Excel show contextual text like *"Chart Area"* when hovering elements, but I found that sometimes they would persist and even reappear in random locations after switching sheets.

This blog documents how I investigated the bug using **VBA, the Win32 API, and WinSpy++**.

---

## Reproducing the Bug

At work the other day, I stumbled upon something odd in Excel. I had a workbook with a bunch of charts, and when I hovered over one of them to see a value and then went to another sheet, the tooltip was still appearing on the other sheet in random locations, and it was only destroyed if I hovered over it or left the Excel grid. I decided to investigate.

At first, I suspected this might be a **resource leak** or even a **use-after-free (UAF)** error.

When I got home this weekend, I got to work trying to debug. I opened up a new workbook and populated it with 49 charts with random data. However, I was unable to reacreate the bug.

I ended up changing Excel's affinity to decrease the number of cores it was running on, figuring my home computer is probably much more powerful than my one at work and that that might have something to do with it. Luckily, that worked.

I then proceeded to record and debug, looking at various window properties via WinSpy++. I found it very interesting to see that the tooltip was not appearing in WinSpy when I refreshed it while it was invisible, but it did appear when it was visible and I refreshed. When it was loaded in WinSpy, even when it was invisible, I could still use the Flash feature to show where it supposedly was on the screen.

Another interesting thing was that the hWnd was the same every time I did the unload/refresh/load/refresh routine. This told me that it was probably not as severe of an error as I had thought, but I was determined to get to the bottom of this.

I inspected other attributes, and the styles TTS_ALWAYSTIP, TTS_NOPREFIX, and WS_EX_TOPMOST were both enabled, which explained why it was showing up on difference sheets.

I made a short tooltip log in VBA to show all the tooltips with their captions and styles, and the only style that changed was the WS_VISIBLE style (=10000000):

| State   |   HWND  | Style  |
|:-------:|:--------|:-------|
|Hidden   |0x27308AA|84000003|
|Visible  |0x27308AA|94000003|

Interestingly enough, all the tooltips had the rightmost digit as 3, which indicates that all had both TTS_ALWAYSTIP and TTP_NOPREFIX. It also showed that there were 64 tooltips always, and looking at WinSpy confirmed this. It seems that these are preloaded, as none of them had a caption except for the ones that I expected to.

Of course, I do not have the source code. However, it appears that this is a harmless, though interesting, UI bug that is causing ghost tooltips. My estimation is that Excel preloads x tooltips, and if it gets overloaded, it might forget to unload one, and its properties get corrupted. This case study shows that much can be done to debug even without source code.

---

## Tools used

- **WinSpy++** to inspect window handles (`HWND`), captions, styles, etc.
- **VBA + WinAPI** calls (`FindWindowEx`, `GetWindowText`, `GetWindowLong`) to enumerate tooltip windows and log their state.
- **Screen recording** to capture the glitch in action.

## What I found

1. **Excel Preallocates Tooltips**
    - Excel creates ~64 tooltip windows (`tooltips_class32`) when it starts.
    - These are not destroyed or recreated; they are reused across the session.

2. **HWND is Stable**
    - Each tooltip keeps the same handle (`HWND`) across sheets and states.
    - No dangling or invalid handles were observed.

3. **Visibility and Caption Change**
    - Style flags (`WS_VISIBLE`) toggle on/off as the tooltip is shown/hidden.
    - Caption text (e.g., `"Chart Area"`) updates dynamically.
    - Sometimes the visibility flag doesn't reset properly, leaving a "ghost tooltip."

---

## Conclusion

This bug is **not a security issue** (no evidence of UAF or memory corruption), but rather a **UI lifecycle bug** in Excel’s tooltip manager. Excel maintains a pool of tooltip windows and sometimes fails to reset one correctly when switching sheets, leading to ghost visuals.

---

## Takeaways

- Debugging Win32 apps is possible even without source code - handles, styles, and captions can reveal a lot.
- What looked at first like a low-level memory bug turned out to be a benign UI issue.  
- Documenting investigations like this is a great way to practice software debugging and systems-level thinking.

## Videos & Code

I’ve uploaded videos showing the bug in action, as well as the VBA logging scripts I used, to [GitHub](https://github.com/trumpetkern27/Excel-Bug---Oct.-2025).

---

## Career Note

I’ve also reported this issue to Microsoft through Excel’s built-in feedback system. As someone learning **C++ and WinAPI** and aiming to become a Software Engineer, I’d love to hear from the Excel team (or others) about this bug and any background they can share.

*If you’ve seen similar issues in other Win32 apps, I’d love to hear your thoughts — drop me a comment on LinkedIn or message me through my [website](the-kern.com)!*