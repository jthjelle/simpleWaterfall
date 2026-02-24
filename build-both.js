const fs = require('fs');
const { execSync } = require('child_process');
const path = require('path');

const pbivizPath = path.join(__dirname, 'pbiviz.json');
const visualTsPath = path.join(__dirname, 'src', 'visual.ts');

const GUID_ORIGINAL = "simpleWaterfall1FE2338C7F9748C2B4633F4B20AACA8A";
const NAME_ORIGINAL = "Waterfall Bridge";

const GUID_TEST = "simpleWaterfallTestB4633F4B20AACA8B";
const NAME_TEST = "Waterfall Bridge (Test)";

// The exact method block we will swap in visual.ts
const LICENSE_CHECK_ORIGINAL = `    private async checkLicensing(): Promise<boolean> {
        // 1. Desktop Check (Free)
        if (this.host.hostEnv === powerbi.common.CustomVisualHostEnv.Desktop) {
            return true;
        }

        // 2. Service Check (AppSource)
        try {
            const licenseInfo = await this.host.licenseManager.getAvailableServicePlans();
            if (!licenseInfo || !licenseInfo.plans) {
                return false;
            }
            
            // Check for any active plan
            const hasActivePlan = licenseInfo.plans.some(plan =>
                plan.state === powerbi.ServicePlanState.Active ||
                plan.state === powerbi.ServicePlanState.Warning
            );
            return hasActivePlan;
            
        } catch (err) {
            console.error('License check failed', err);
            return false;
        }
    }`;

const LICENSE_CHECK_TEST = `    private async checkLicensing(): Promise<boolean> {
        // ALWAYS RETURN TRUE FOR TEST VERSION (NO LICENSE CHECK)
        return true;
    }`;

function setPbiviz(guid, name) {
    const data = JSON.parse(fs.readFileSync(pbivizPath, 'utf8'));
    data.visual.guid = guid;
    data.visual.displayName = name;
    fs.writeFileSync(pbivizPath, JSON.stringify(data, null, 4));
}

function setVisualTs(isTest) {
    let content = fs.readFileSync(visualTsPath, 'utf8');

    // Find the start of the method
    const startRegex = /private\s+async\s+checkLicensing\(\):\s*Promise<boolean>\s*\{/;
    const match = content.match(startRegex);
    if (!match) throw new Error("Could not find checkLicensing method signature in visual.ts");

    const startIndex = match.index;

    // Find the matching closing brace for this method
    let braceCount = 0;
    let endIndex = -1;
    let started = false;

    for (let i = startIndex; i < content.length; i++) {
        if (content[i] === '{') {
            braceCount++;
            started = true;
        } else if (content[i] === '}') {
            braceCount--;
        }

        if (started && braceCount === 0) {
            endIndex = i + 1;
            break;
        }
    }

    if (endIndex === -1) throw new Error("Could not find closing brace for checkLicensing");

    const originalBlock = content.substring(startIndex, endIndex);

    let replacement = isTest ? LICENSE_CHECK_TEST : LICENSE_CHECK_ORIGINAL;

    let newContent = content.substring(0, startIndex) + replacement + content.substring(endIndex);
    fs.writeFileSync(visualTsPath, newContent);
}

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
    fs.renameSync(output, target);
    console.log(`Saved as dist/${packageName}.pbiviz`);
}

console.log("==========================================");
console.log("   DUAL BUILD SCRIPT FOR POWERBI VISUAL   ");
console.log("==========================================\\n");

try {
    // 1. BUILD TEST OFFER
    console.log(">>> [1/2] PREPARING TEST VERSION (license check removed) <<<");
    setPbiviz(GUID_TEST, NAME_TEST);
    setVisualTs(true);
    build('simpleWaterfall_Test_NoLicense');

    console.log("\\n>>> [2/2] PREPARING ORIGINAL VERSION (official AppSource config) <<<");
    // 2. BUILD ORIGINAL OFFER
    setPbiviz(GUID_ORIGINAL, NAME_ORIGINAL);
    setVisualTs(false);
    build('simpleWaterfall_Original');

    // Copy the original back to standard name just in case local host server needs it
    fs.copyFileSync(
        path.join(__dirname, 'dist', 'simpleWaterfall_Original.pbiviz'),
        path.join(__dirname, 'dist', 'simpleWaterfall.pbiviz')
    );

    console.log("\\n✅ SUCCESS: Both packages built into the dist/ directory!");
} catch (e) {
    console.error("\\n❌ Build failed:", e.message);
} finally {
    // Always attempt recovery to Original State so developer can keep working
    console.log("Restoring project to original code state...");
    try {
        setPbiviz(GUID_ORIGINAL, NAME_ORIGINAL);
        setVisualTs(false);
    } catch (recoverErr) {
        console.error("Warning: Could not restore to original code state:", recoverErr.message);
    }
}
