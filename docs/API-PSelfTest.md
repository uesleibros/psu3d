# PSelfTest

**Self test**

Exercises every module of the library against known answers and reports what passed. Meant to be the first thing run after importing the engine into a project, because the failures worth catching there are silent ones: a module left out of the import, a class that did not come across, or the JSON reader missing.

> Runs entirely in memory. The renderer is put in dry run so the whole pipeline executes without a slide, which means this can be run from the VBE with no presentation open and leaves nothing behind.

> **Scope.** Deliberately NOT Option Private Module, for the same reason as PDemo: the macro dialog only lists public entry points, and this one has to be startable from there. Every engine module stays private.

## Index

**Entry points.** [`Psu3DSelfTest`](#psu3dselftest), [`RunAll`](#runall)

## Members

### Psu3DSelfTest

```vba
Public Sub Psu3DSelfTest()
```

Runs every check and shows the result.

> The one call to make after importing the library.

### RunAll

```vba
Public Function RunAll() As String
```

Runs every check and returns the report.

**Returns.** A line per failure, then a summary line.
