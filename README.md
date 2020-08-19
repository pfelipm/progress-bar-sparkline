![Banner progress-bar-sparkline](https://docs.google.com/drawings/d/e/2PACX-1vQmkaz4vcDu-bqPwiKPugfWiCQdE1es9SSeM2x4MAk-6sFRG2nSFQKfjvAxpoMmKKBUSLivl8wcQbzy/pub?w=1280&h=320)

# Progress bar with SPARKLINES for Google Sheets

Extremely simple functions to create and manage something akin to progress bars in Google Sheets using the built in function `SPARKLINE()`, built just for my own learning of Apps Script libraries :blush: . Two versions are provided:

*   **PB.gs**: Implemented using a simple constructor function. This provides JSDoc contextual help (not for methods, though) when importing as a GAS library.
*   **PB (ES6 class).gs**: Implemented using a ES6 class. I've not managed to get any JSDoc-style help at all. In this case the class needs to be explicitly exported for it to be used as a library like this:

```
// This makes class ProgressBarClassExt methods available when using as a library (thanks to @stevenbazyl for the tip)
var ProgressBarES6Ext = ProgressBarES6; 
```
