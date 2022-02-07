interface Array<T> {
  unique(): T[];
}

Array.prototype.unique = function () {
  var arr = [];
  for (let i = 0; i < this.length; i++) {
    if (arr.indexOf(this[i]) == -1) {
      arr.push(this[i]);
    }
  }
  return arr;
};
