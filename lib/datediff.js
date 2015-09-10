// https://gist.github.com/remino/1563963

function DateDiff(date1, date2) {
  this.days = null;
  this.hours = null;
  this.minutes = null;
  this.seconds = null;
  this.date1 = date1;
  this.date2 = date2;

  this.init();
}

DateDiff.prototype.init = function() {
  var data = new DateMeasure(this.date1 - this.date2);
  this.days = data.days;
  this.hours = data.hours;
  this.minutes = data.minutes;
  this.seconds = data.seconds;
};

function DateMeasure(ms) {
  var d, h, m, s;
  s = Math.floor(ms / 1000);
  m = Math.floor(s / 60);
  s = s % 60;
  h = Math.floor(m / 60);
  m = m % 60;
  d = Math.floor(h / 24);
  h = h % 24;
  
  this.days = d;
  this.hours = h;
  this.minutes = m;
  this.seconds = s;
};

Date.diff = function(date1, date2) {
  return new DateDiff(date1, date2);
};

Date.prototype.diff = function(date2) {
  return new DateDiff(this, date2);
};
