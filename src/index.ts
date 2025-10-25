interface Person {
  name: string;
  age?: number;
}

class Greeter {
  private greeting: string;

  constructor(message: string) {
    this.greeting = message;
  }

  greet(person: Person): string {
    return `${this.greeting}, ${person.name}!`;
  }
}

const greeter = new Greeter("Hello");
const user: Person = { name: "TypeScript World", age: 5 };

console.log(greeter.greet(user));