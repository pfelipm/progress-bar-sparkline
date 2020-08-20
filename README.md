![Banner progress-bar-sparkline](https://docs.google.com/drawings/d/e/2PACX-1vQmkaz4vcDu-bqPwiKPugfWiCQdE1es9SSeM2x4MAk-6sFRG2nSFQKfjvAxpoMmKKBUSLivl8wcQbzy/pub?w=1280&h=320)

# Progress bar with SPARKLINES for Google Sheets

Extremely simple functions to create and manage something akin to progress bars in Google Sheets using the built in function `SPARKLINE()`, built just for my own learning of Apps Script libraries :blush: . Two versions are provided:

*   **PB.gs**: Implementation using a [simple constructor function](https://developer.mozilla.org/es/docs/Learn/JavaScript/Objects/Object-oriented_JS). This provides JSDoc contextual help (just for the constructor, though) when importing as a GAS library.
*   **PB (ES6 class).gs**: Implementation using a [ES6 class](https://github.com/DrkSephy/es6-cheatsheet#classes). I've not managed to get any JSDoc-style help at all. In this case the class needs to be explicitly exported for it to be used as a library.

Further context & motivation in this post: [Barras de progreso Apps Script usando SPARKLINE()](https://pablofelip.online/barras-progreso-apps-script-sparkline/).

# Instructions

Two ways to use this:

1.  Open the Apps Script editor in your spreadsheet (`Tools` ⇒ `Script editor`), paste the provided code (**PB.gs** and/or **PB (ES6 class).gs**) and save. You must use the _not-so-new-now_ JavaScript V8 GAS engine (`Ejecutar` ⇒ `Enable new Apps Script runtime ... V8`).
2.  Import as library:
    *   Open your own project in the GAS editor.
    *   Resources ⇒ Libraries.
    *   Add a library, project key: **Mvzrd2GnnBN6AmaTRJzQGHIlk-AYma6-o**.
    *   Input a suitable identifier, e.g. **pbs** (don't ever use 'pb' or you won't get contextual help for the constructor :man\_shrugging:).
    *   Save changes.

![Selección_374](https://user-images.githubusercontent.com/12829262/90753459-ee79e280-e2d8-11ea-9bbf-b46605bc521b.png)

See demo and code sample here :point\_right: [Progress bar SPARKLINE # demo](https://docs.google.com/spreadsheets/d/1NYzgkpvAhWJdldczHv4EgRfznpjeJ_lRDrkPLGy73iQ/template/preview) :point\_left:.

![Progress bar SPARKLINE # demo - Hojas de cálculo de Google](https://pablofelip.online/media/posts/14/ezgif.com-video-to-gif.gif)

In your GAS project, initialize a progress bar in cell **A10** of sheet **Test** with an initial value of **0** and max value of **100** like this (notice the `pb.`, I am using the constructor function version in **PB.gs** as an imported library here). This will inject a SPARKLINE() function in cell `Test!A10`:

```javascript
let progressBar1 = new pbs.ProgressBar('Test!A10', 0, 100);
```

Unspecified parameters (_bar colors_, _step_, _flush sheet_ will take defaults). `Value`, `Max` and step (`reDrawEvery`) parameters are always rounded.

To update progress:

```javascript
progressBar1.value = 25;
progressBar1.update();
```

...or simply:

```javascript
progressBar1.update(25);
```

You can also do:

```javascript
progressBar1.clear();  // Set at 0%
progressBar1.halve();  // Set at 50%
progressBar1.fill();   // Set at 100%
```

You can instantiate more progress bars, should you need it, and manage each of them separately.

Check source code for all class properties and more info about methods.

# Under the hood (code review :gear:)

Nothing really relevant this time besides (a) the need to (somewhat) export ES6 classes when using them inside a library just to be able to invoke their methods and (b) the impossibility (to the best of my knowledge) to prepare proper JSDoc contextual help. A pity, although I may be missing something.

```javascript
// This makes class ProgressBarClassExt methods available when used as a library (thanks to @stevenbazyl for the tip)
var ProgressBarES6Ext = ProgressBarES6;
```

I think that in order to prevent non-integer numbers in some properties more elegantly I should have used [setter functions](https://developer.mozilla.org/es/docs/Web/JavaScript/Referencia/Funciones/set) inside my class... but anyway, in the end this is little more than a first attempt at getting a grasp of libraries.

# License

© 2020 Pablo Felip Monferrer ([@pfelipm](https://twitter.com/pfelipm)). Provided under MIT License.
