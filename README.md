![Banner progress-bar-sparkline](https://docs.google.com/drawings/d/e/2PACX-1vQmkaz4vcDu-bqPwiKPugfWiCQdE1es9SSeM2x4MAk-6sFRG2nSFQKfjvAxpoMmKKBUSLivl8wcQbzy/pub?w=1280&h=320)

# Progress bar with SPARKLINES for Google Sheets

Extremely simple functions to create and manage something akin to progress bars in Google Sheets using the built in function `SPARKLINE()`, built just for my own learning of Apps Script libraries :blush: . Two versions are provided:

*   **PB.gs**: Implemented using a simple constructor function. This provides JSDoc contextual help (not for methods, though) when importing as a GAS library.
*   **PB (ES6 class).gs**: Implemented using a ES6 class. I've not managed to get any JSDoc-style help at all. In this case the class needs to be explicitly exported for it to be used as a library like this:

```
// This makes class ProgressBarClassExt methods available when using as a library (thanks to @stevenbazyl for the tip)
var ProgressBarES6Ext = ProgressBarES6; 
```

# **Modo de uso**

Two ways:

1.  Open the Apps Script editor in your spreadsheet (`Tools` 🠲 `Script editor`), paste the provided code (**PB.gs** and/or **PB (ES6 class).gs**) and save. You must use the _not-so-new-now_ JavaScript V8 GAS engine (`Ejecutar` 🠲 `Enable new Apps Script runtime ... V8`).
2.  Import as library:
    *   Open GAS editor.
    *   Resources 🠲 Libraries.
    *   Add a library, project key: **Mvzrd2GnnBN6AmaTRJzQGHIlk-AYma6-o**.
    *   Input a suitable identifier (e.g. **pb**).
    *   Save changes.

![addaslib](https://user-images.githubusercontent.com/12829262/90613499-be174300-e209-11ea-9ee4-2da9cee2357c.png)

See demo and here :point\_right: [Progress bar SPARKLINE # demo](https://docs.google.com/spreadsheets/d/1NYzgkpvAhWJdldczHv4EgRfznpjeJ_lRDrkPLGy73iQ/template/preview) :point\_left:.

![Progress bar SPARKLINE # demo - Hojas de cálculo de Google](https://user-images.githubusercontent.com/12829262/90616041-36cbce80-e20d-11ea-99a0-7973273da378.gif)

In your project, initialize a progress bar in cell **A10** of sheet **Test** with an initial value of **0** and max value of **100** like this:

```
let progressBar1 = new pb.ProgressBar('Test!A10', 0, 100);
```

Other parameters (bar colors, step, flush sheet will take defaults).

Check source code (demo & library) for class properties and methods.

# **Licencia**

© 2020 Pablo Felip Monferrer ([@pfelipm](https://twitter.com/pfelipm)). Se distribuye bajo licencia MIT.
