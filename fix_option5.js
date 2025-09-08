const fs = require('fs');

// Read the script file
let content = fs.readFileSync('script.js', 'utf8');

// Fix the syntax error and add Option 5 properly
const brokenPattern = /            } else if \(version === 5\) \{\s*\/\/ Option 5: Duplicate of Option 2 - always starts in test mode\s*this\.isTestMode = true;\s*localStorage\.setItem\("testMode", "true"\);\s*this\.showTestModeToggle\(\);\s*this\.setupTestModeToggle\(\);\s*this\.hideTestModePillOption4\(\);\s*this\.greyOutRows11To20\(\);\s*this\.showTestModeToastOnLoad\(\); \/\/ Show toast when switching to Option 5/;

const fixedOption5 = `        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }`;

// Replace the broken pattern
content = content.replace(brokenPattern, fixedOption5);

// Write back to file
fs.writeFileSync('script.js', content);
console.log('Option 5 fixed successfully!');
