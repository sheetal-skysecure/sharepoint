const fs = require('fs');
const content = fs.readFileSync('c:\\Users\\SheetalSinha\\skysecureinternal\\spfx-learning-center\\src\\webparts\\adminAccess\\components\\AdminPortal.tsx', 'utf8');
const lines = content.split('\n');
for (let i = 2120; i < 2140; i++) {
    process.stdout.write((i + 1) + ': ' + JSON.stringify(lines[i]) + '\n');
}
