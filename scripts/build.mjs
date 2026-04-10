import { cpSync, existsSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const root = process.cwd();
const dist = join(root, 'dist');

if (existsSync(dist)) rmSync(dist, { recursive: true, force: true });
mkdirSync(dist, { recursive: true });

for (const file of ['index.html', 'app.js', 'styles.css']) {
  cpSync(join(root, file), join(dist, file));
}

cpSync(join(root, 'README.md'), join(dist, 'README.md'));
cpSync(join(root, 'README.txt'), join(dist, 'README.txt'));
cpSync(join(root, '404.html'), join(dist, '404.html'));
writeFileSync(join(dist, '.nojekyll'), '');

console.log('Build complete: dist/');
