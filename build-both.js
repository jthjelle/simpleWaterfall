const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

const pbivizPath = path.join(__dirname, 'pbiviz.json');

function build(packageName) {
    console.log(`Building ${packageName}...`);
    execSync('npx pbiviz package', { stdio: 'inherit', cwd: __dirname });

    const data = JSON.parse(fs.readFileSync(pbivizPath, 'utf8'));
    const outputName = `${data.visual.guid}.${data.visual.version}.pbiviz`;
    const output = path.join(__dirname, 'dist', outputName);

    const target = path.join(__dirname, 'dist', `${packageName}.pbiviz`);
    if (fs.existsSync(target)) {
        fs.unlinkSync(target);
    }

    fs.copyFileSync(output, target);
    fs.copyFileSync(output, path.join(__dirname, 'dist', 'simpleWaterfall.pbiviz'));
    console.log(`Saved as dist/${packageName}.pbiviz`);
}

console.log("==========================================");
console.log("   FREE BUILD SCRIPT FOR POWERBI VISUAL   ");
console.log("==========================================\\n");

try {
    build('simpleWaterfall_NoLicense');
    console.log("\\nSUCCESS: Free packages built into the dist/ directory!");
} catch (e) {
    console.error("\\nBuild failed:", e.message);
    process.exitCode = 1;
}
