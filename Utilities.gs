function load_DateUtilities() {
  return {
    fixDate: function(dte){
      
      // If dte is a number then it contains a date in the form of the number of days since
      // 00:00 on 12/31/1899 in the current time zone
      
      if (typeof dte == "number"){
        return dte
      }
      
      // If dte is a date object then the getTime property contains the date as the number of
      // milliseconds since 00:00 on January 1, 1970 in GMT.  This value is converted to be
      // compatible with the other format
      
      return (dte.getTime() / 86400000) + 25569 - dte.getTimezoneOffset() / (60 * 24)
    }, 
    inDays: function(d1, d2) {
      var t2 = d2.getTime();
      var t1 = d1.getTime();
      
      return parseInt((t2-t1)/(24*3600*1000));
    },
    inWeeks: function(d1, d2) {
      var t2 = d2.getTime();
      var t1 = d1.getTime();
      
      return parseInt((t2-t1)/(24*3600*1000*7));
    },
    inMonths: function(d1, d2) {
      var d1Y = d1.getFullYear();
      var d2Y = d2.getFullYear();
      var d1M = d1.getMonth();
      var d2M = d2.getMonth();
      
      return (d2M+12*d2Y)-(d1M+12*d1Y);
    },
    inYears: function(d1, d2) {
      return d2.getFullYear()-d1.getFullYear();
    },
    isWeekend: function(d) {
      var day = d.getDay();
      return (day == 6) || (day == 0);
    },
    isMonday: function(d) {
      var day = d.getDay();
      return day == 1;
    },
    isSaturday: function(d) {
      var day = d.getDay();
      return day == 6;
    },
    isSunday: function(d) {
      var day = d.getDay();
      return day == 0;
    },
    numberOfWeekendsBetween: function(d1, d2) {
      var ndays = 1 + Math.round((d2.getTime()-d1.getTime())/(24*3600*1000));
      var nsaturdays = Math.floor( (d1.getDay()+ndays) / 7 );
      return 2*nsaturdays + (d1.getDay()==0) - (d2.getDay()==6);
    },
    addDays: function(d, n) {
      d.setTime( d.getTime() + n * 86400000 );
      return d;
    },
    toString: function(d) {
      var yyyy = d.getFullYear().toString();
      var mm = (d.getMonth()+1).toString(); // getMonth() is zero-based
      var dd  = d.getDate().toString();
      return yyyy + (mm.length===2?mm:"0"+mm[0]) + (dd.length===2?dd:"0"+dd[0]); // padding
    }
  }
}

function hexToRgb(hex) {
    // Expand shorthand form (e.g. "03F") to full form (e.g. "0033FF")
    var shorthandRegex = /^#?([a-f\d])([a-f\d])([a-f\d])$/i;
    hex = hex.replace(shorthandRegex, function(m, r, g, b) {
        return r + r + g + g + b + b;
    });

    var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
    } : null;
}
