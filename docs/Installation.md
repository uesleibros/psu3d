# Installation

## Requirements

PowerPoint with VBA enabled, and a file saved as `.pptm`. There is nothing else to install: no DLL, no reference to tick in Tools, no setup program.

The library uses a handful of Windows API functions, declared inside the modules that use them: `QueryPerformanceCounter` for the clock, and `GetAsyncKeyState`, `GetCursorPos`, `SetCursorPos` and `ShowCursor` for the demo's input. All of them ship with Windows.

## Importing

1. Open the VBE with `Alt+F11`.
2. **File**, then **Import File**, or `Ctrl+M`.
3. Import the files from `engine/`.

Import order does not matter to the compiler, only to a reader.

```
PCore.bas        types, maths, colour, clock
PLighting.bas    light and fog
PMaterial.cls    one surface
PMaterials.bas   material registry
PCanvas.cls      where to project
PCamera.cls      where to look from
PRenderer.cls    how to draw
PScene.cls       what exists
PBody.cls        what moves and collides       (optional)
PLevel.bas       read and write JSON           (optional)
Psu3D.bas        facade
PSelfTest.bas    self test                     (optional)
PDemo.bas        playable example level        (optional)
```

## The minimum

If you only want to draw in 3D, eight modules are enough:

`PCore`, `PLighting`, `PMaterial`, `PMaterials`, `PCanvas`, `PCamera`, `PRenderer`, `PScene`.

`Psu3D.bas` is a convenience layer over those eight and costs nothing to bring along.

## The one external module

Psu3D bundles no third party code, and the core calls nothing outside itself. VBA compiles the whole project at once, so importing a module without what it calls fails to compile. That is why the core calls nothing: all twelve modules other than `PLevel` stand on their own.

`PLevel.bas` needs a JSON parser, and uses it throughout. Take it from **https://github.com/vbacollective/json** and import `JSON.cls`.

If you are not reading or writing levels from files, do not import `PLevel.bas` and the question disappears.

Mouse and keyboard are not a dependency. `PDemo` declares the Windows API functions it needs at the top of the file, next to the ones for the keyboard.

## Checking the install

Run `PSelfTest.Psu3DSelfTest`. It is 178 assertions against known answers, running entirely in memory: the renderer is put into `DryRun`, so the whole pipeline executes with no slide and leaves nothing behind. It can be run from the VBE with no presentation open.

It exists because the failures that matter in a hand imported library are silent ones: a module left out, a class that did not come across, a parser that is missing. The report names every check that failed.

One detail about `.cls` files: when importing a class the VBE reads the name from the `VB_Name` attribute at the top of the file, not from the file name. Renaming the file does not rename the class. Do not rename them.

## Enabling macros

PowerPoint blocks macros by default in a file downloaded from the internet. Right click the `.pptm`, choose **Properties**, and tick **Unblock** at the bottom of the General tab. Without that the VBE opens but nothing runs.
