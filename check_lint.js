const build = require('@microsoft/sp-build-web');
console.log('lintCmd:', typeof build.lintCmd);
console.log('eslintCmd:', typeof build.eslintCmd);
console.log('eslint:', typeof build.eslint);
console.log('tslint:', typeof build.tslint);
console.log('tslintCmd:', typeof build.tslintCmd);
