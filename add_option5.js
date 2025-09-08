const fs = require('fs');

// Read the script file
let content = fs.readFileSync('script.js', 'utf8');

// Find the end of Option 4 and add Option 5
const option4End = '            } else {\n                this.enableAllRows();\n            }\n        }\n    }';
const option5Code = '            } else {\n                this.enableAllRows();\n            }\n        } else if (version === 5) {\n            // Option 5: Duplicate of Option 2 - always starts in test mode\n            this.isTestMode = true;\n            localStorage.setItem("testMode", "true");\n            this.showTestModeToggle();\n            this.setupTestModeToggle();\n            this.hideTestModePillOption4();\n            this.greyOutRows11To20();\n            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5\n        }\n    }';

content = content.replace(option4End, option5Code);

// Write back to file
fs.writeFileSync('script.js', content);
console.log('Option 5 added successfully!');
