const fs = require('fs');
const vm = require('vm');

try {
  const html = fs.readFileSync('index.html', 'utf8');
  // Extract the JS block (everything between the last <script> and </script>)
  const scriptRegex = /<script>([\s\S]*?)<\/script>/g;
  let match;
  let scriptCount = 0;
  
  while ((match = scriptRegex.exec(html)) !== null) {
    scriptCount++;
    const jsCode = match[1];
    console.log(`Checking script block ${scriptCount}...`);
    try {
      new vm.Script(jsCode);
      console.log(`Script block ${scriptCount} is syntactically correct.`);
    } catch (syntaxError) {
      console.error(`Syntax error in script block ${scriptCount}:`);
      console.error(syntaxError.message);
      // Print the line numbers around the error
      const lines = jsCode.split('\n');
      const errorLine = syntaxError.stack.split('\n')[0].match(/:(\d+)/);
      if (errorLine) {
        const lineNum = parseInt(errorLine[1], 10);
        console.error(`Error line (${lineNum}): ${lines[lineNum - 1]}`);
        console.error(`Context:`);
        for (let i = Math.max(0, lineNum - 5); i < Math.min(lines.length, lineNum + 5); i++) {
          console.error(`${i + 1}: ${lines[i]}`);
        }
      }
    }
  }
} catch (err) {
  console.error('File read error:', err);
}
