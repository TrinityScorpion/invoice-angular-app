const fs = require('fs');
const { join } = require('path');
const { promisify } = require('util');

const rmdir = promisify(fs.rmdir);
const lstat = promisify(fs.lstat);
const unlink = promisify(fs.unlink);
const readdir = promisify(fs.readdir);

const rm = async dir => {
  const stat = await lstat(dir);
  if (!stat.isDirectory()) {
    await unlink(dir);
    return;
  }
  const files = await readdir(dir);
  for (const file of files) {
    await rm(join(dir, file));
  }
  await rmdir(dir);
};

module.exports = rm;