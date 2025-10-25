import { Demand } from './demand.mjs'; // Note the .mjs extension
const instance = new Demand("Samir I. Sheth", './data/Picture1.jpg');
console.log(instance.sayHello()); // Output: Hello, my name is Alice!
instance.makeLetter();
