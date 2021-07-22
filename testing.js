const date1 = new Date('2019-03-07');
const date2 = new Date('2019-03-05');
const diffTime = Math.abs(date2 - date1);
const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
console.log(diffTime + " milliseconds");
console.log(diffDays + " days");