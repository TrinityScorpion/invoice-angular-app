## tiny-file

> Tiny file utils for Node.js

[![tiny-file](https://img.shields.io/npm/v/tiny-file.svg)](https://npmjs.org/tiny-file)

### Installation

```bash
$ npm install tiny-file
```

### Example

```js
const { mkdir, rmdir } = require('tiny-file');

(async function() {
  await mkdir('A/B/C/D');
  await rmdir('A');
  console.log('Done!');
})();

```

### Contributing
- Fork this Repo first
- Clone your Repo
- Install dependencies by `$ npm install`
- Checkout a feature branch
- Feel free to add your features
- Make sure your features are fully tested
- Publish your local branch, Open a pull request
- Enjoy hacking <3

### MIT

This work is licensed under the [MIT license](./LICENSE).

---