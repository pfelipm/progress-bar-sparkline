/**
 * A simple constructor function to implement a progress bar style indicator using the SPARKLINE() function in Google Sheets
 *
 * @constructor
 *
 * @param {string} cellFullRef Full cell reference (sheet + cell reference, e.g. 'Sheet 1!A1'.
 * @param {number} value Actual value for the progress bar.
 * @param {number} max Maximum value for the progress bar.
 * @param {number} reDrawEvery Progress bar will only redraw every reDrawEvery calls.
 * @param {string} color1 String representing the hexa color of the filled part (left) of the progress bar.
 * @param {string} color2 String representing the hexa color of the empty part (right) of the progress bar.
 * @param {boolean} flushSheet If set to TRUE forces a flush of the sheet after every update.
 *
 *  Contextual JSDoc help is not working for methods when using as a library.
 *
 * Properties:
 * -----------
 * cell >> cel reference.
 * value >> current value.
 * max >> max value.
 * reDrawEvery >> Redraw steps.
 * reDrawCount >> Actual redraw step.
 * color1 >> left color (full).
 * color2 >> right color (empty).
 * flush >> Force flush behaviour.
 *
 * Methods:
 * --------
 * update([value], forceRedraw) >> Updates progress bar status. If value is not provided property .value is used.
 * clear() >> Sets progress bar to 0%.
 * fill()  >> Sets progress bar to 100%.
 * halve()  >> Sets progress bar to 50%.
 *
 * @OnlyCurrentDoc
 *
 * MIT License
 * Copyright (c) 2020 Pablo Felip Monferrer(@pfelipm)
 */ 
 
function ProgressBar(cellFullRef, value = 0, max = 100, reDrawEvery = 1, color1 = '#46bdc6', color2 = '#999999', flushSheet = false) {
  
  this.cell = cellFullRef;
  this.value = value;
  this.max = max;
  this.reDrawEvery = reDrawEvery;
  this.reDrawCount = -1; // First update inside constructor does not count
  this.color1 = color1;
  this.color2 = color2;
  this.flush = flushSheet;
                
  //  Updates progress bar, uses object properties if a new value is not provided
  
  /**
   * Updates progress bar, uses object properties if a new value is not provided
   * @param {number} value Actual value for the progress bar.
   * @param {boolean} forceRedraw Redraws progress bar even though reDrawcount is not reached.
   */
  
  this.update = function(value = this.value, forceRedraw) {
    
    this.reDrawCount = (this.reDrawCount + 1) % this.reDrawEvery;
    if (this.reDrawCount == 0 || forceRedraw) {
    
      this.value = value < 0 ? 0 : value > this.max ? this.max : value;
      SpreadsheetApp.getActiveSpreadsheet().getRange(this.cell).setFormula('SPARKLINE({' + this.value
                                                                                         + '\\' 
                                                                                         + (this.max - this.value) 
                                                                                         + '};{"charttype"\\"bar";"color1"\\"' 
                                                                                         + this.color1 
                                                                                         + '";"color2"\\"' 
                                                                                         + this.color2 + '"})');
      // Using ES6 strings in this line 👇 breaks auto indenting in the editor
      // SpreadsheetApp.getActiveSpreadsheet().getRange(this.cell).setFormula(`SPARKLINE({${this.value}\\${this.max - this.value}};{"charttype"\\"bar";"color1"\\"${this.color1}";"color2"\\"${this.color2}"})`);                                                                              
      
      if (this.flush) SpreadsheetApp.flush();
    
    }
                  
  }

  // Sets progress bar to 0%                  
  this.clear = function() {this.update(0, true);}
  
  // Sets progress bar to 100%                  
  this.fill = function() {this.update(this.max, true);}
                  
  // Sets progress bar to 50%
  this.halve = function() {this.update(this.max / 2, true);}
  
  // Draw progress bar after object constructor code + methods (function constructors are not hoisted)
  this.update(this.value, true);
  
}