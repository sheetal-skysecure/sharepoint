const build = require('@microsoft/sp-build-web');
const keys = Object.keys(build);
keys.forEach(key => {
    try {
        const val = build[key];
        if (val && typeof val === 'object' && val.enabled !== undefined) {
            console.log(`${key} has enabled property: ${val.enabled}`);
        } else {
            console.log(`${key} exists`);
        }
    } catch (e) {
        console.log(`${key} error accessing`);
    }
});
