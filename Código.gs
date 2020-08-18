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
 * @param {boolean} flushSheet If set to TRUE forces a flush of the sheet with every update.
 *
 * Properties:
 * -----------
 * cell >> cel reference.
 * value >> current value.
 * max >> max value.
 * reDrawEvery >> Redraw step.
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
 * half()  >> Sets progress bar to 50%.
 *
 * @OnlyCurrentDoc
 */
 
function ProgressBar(cellFullRef, value = 0, max = 100, reDrawEvery = 1, color1 = '#46bdc6', color2 = '#999999', flushSheet = false) {
  
  this.cell = cellFullRef;
  this.value = value;
  this.max = max;
  this.reDrawEvery = reDrawEvery;
  this.reDrawCount = 0;
  this.color1 = color1;
  this.color2 = color2;
  this.flush = flushSheet;
                
  //  Updates progress bar, uses object properties if a new value is not provided
  this.update = function(value = this.value, forceRedraw) {
    
    this.reDrawCount = (this.reDrawCount + 1) % this.reDrawEvery
    if (this.reDrawCount == 0 || forceRedraw) {
    
      this.value = value < 0 ? 0 : value > this.max ? this.max : value;
      SpreadsheetApp.getActiveSpreadsheet().getRange(this.cell).setFormula('SPARKLINE({' + this.value
                                                                                         + '\\' 
                                                                                         + (this.max - this.value) 
                                                                                         + '};{"charttype"\\"bar";"color1"\\"' 
                                                                                         + this.color1 
                                                                                         + '";"color2"\\"' 
                                                                                         + this.color2 + '"})');
      // SpreadsheetApp.getActiveSpreadsheet().getRange(this.cell).setFormula(`SPARKLINE({${this.value}\\${this.max - this.value}};{"charttype"\\"bar";"color1"\\"${this.color1}";"color2"\\"${this.color2}"})`);                                                                              
      if (this.flush) SpreadsheetApp.flush();
    
    }
                  
  }

  // Resets progress bar                  
  this.clear = function() {this.update(0, true);}
  
  // Fills progress bar                  
  this.fill = function() {this.update(this.max, true);}
                  
  // Fills progress bar                  
  this.half = function() {this.update(Math.round(this.max / 2), true);}
  
  // Draw progress bar ends object constructor (function constructors are not hoisted)
  
  this.update(this.value, true);
  
}