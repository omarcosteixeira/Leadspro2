import fs from 'fs';
let rules = fs.readFileSync('firestore.rules', 'utf8');

const matchStart = '    match /artifacts/gestaodeleadspro-d4230/';
let startIndex = rules.indexOf(matchStart);

if (startIndex > -1) {
    let block = rules.substring(startIndex, rules.lastIndexOf('  }\n}'));
    let gestaoproBlock = block.replace(/gestaodeleadspro-d4230/g, 'gestaopro-761e1');
    
    // Duplicate the block
    let newRules = rules.substring(0, startIndex) + gestaoproBlock + '\n' + block + '\n  }\n}';
    fs.writeFileSync('firestore.rules', newRules);
}
