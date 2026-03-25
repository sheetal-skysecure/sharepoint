const build = require('@microsoft/sp-build-web');
console.log(Object.keys(build));
if (build.eslint) console.log('eslint exists');
if (build.eslintCmd) console.log('eslintCmd exists');
if (build.tslintCmd) console.log('tslintCmd exists');
