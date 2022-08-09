const fs = require('fs');
const path = require('path');
const { promisify } = require('util');

const mkdir = promisify(fs.mkdir);

const mkdirp = async dir => {
  const parent = path.dirname(dir);
  if (parent === dir) return;

  await mkdirp(parent);

  try {
    await mkdir(dir);
  } catch (err) {
    if (err.code !== 'EEXIST') throw err;
  }
};

module.exports = mkdirp;