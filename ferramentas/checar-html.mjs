/* Confere o JavaScript embutido nos .html do repo.
 *
 * Existe porque `node --check` não basta: ele valida SINTAXE, e o erro que já
 * derrubou o dashboard quatro vezes é outro — um símbolo que some junto quando
 * eu removo um bloco vizinho. `taxaEm`, `amostraNecessaria`, `URL_QUIZ` e
 * `copiarUrl` saíram assim, cada uma em código perfeitamente válido, e só
 * apareceram como "Erro ao carregar" na tela.
 *
 * Então roda no-undef do ESLint, que é exatamente essa checagem.
 *
 *   node ferramentas/checar-html.mjs [arquivo.html ...]
 *   node ferramentas/checar-html.mjs --autoteste
 */
import { readFileSync, writeFileSync, rmSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

const GLOBAIS = [
  'window', 'document', 'console', 'location', 'navigator', 'history',
  'innerWidth', 'innerHeight', 'getSelection', 'getComputedStyle', 'matchMedia',
  'setTimeout', 'clearTimeout', 'setInterval', 'clearInterval',
  'requestAnimationFrame', 'localStorage', 'sessionStorage', 'fetch', 'crypto',
  'URLSearchParams', 'URL', 'alert', 'confirm', 'prompt', 'structuredClone',
  'Chart', 'supabase',       // vêm de <script src> e não existem neste recorte
  'fbq', 'dataLayer', 'gtag',// pixel do Meta e GTM, injetados pelo snippet deles
  'Image', 'FormData', 'Blob', 'IntersectionObserver', 'MutationObserver',
];

/* fileURLToPath e não .pathname: o caminho tem espaço ('pagina de vendas'),
   que numa URL vem como %20 e aponta pra uma pasta que não existe. */
const raiz = fileURLToPath(new URL('..', import.meta.url));

const cfg = join(mkdtempSync(join(tmpdir(), 'checar-')), 'eslint.config.mjs');
writeFileSync(cfg,
  'export default [{languageOptions:{ecmaVersion:2022,sourceType:"script",globals:' +
  JSON.stringify(Object.fromEntries(GLOBAIS.map((g) => [g, 'readonly']))) +
  '},rules:{"no-undef":"error"}}];\n');

/* O .js temporário vive DENTRO do repo. Fora dele o ESLint 9 devolve zero
   resultados sem reclamar de nada, e isso é indistinguível de "passou" — foi
   exatamente assim que a primeira versão deste script aprovou um arquivo com
   URL_QUIZ faltando. */
const tmpJs = join(raiz, '.checar-tmp.js');

/* Devolve a lista de mensagens do ESLint, ou null se ele não rodou. Separar os
   dois é o ponto: "não sei dizer" e "está limpo" não podem virar a mesma coisa. */
function lint(codigo) {
  writeFileSync(tmpJs, codigo);
  let saida;
  try {
    saida = execFileSync('npx', ['--no-install', 'eslint', '--no-ignore', '-f', 'json',
                                 '--config', cfg, tmpJs],
                         { cwd: raiz, encoding: 'utf8', stdio: 'pipe' });
  } catch (e) {
    saida = String(e.stdout ?? '');            // exit 1 = achou erro, e o JSON vem no stdout
    if (!saida.trim()) return null;            // exit sem saída = não rodou
  }
  let res;
  try { res = JSON.parse(saida); } catch { return null; }
  return Array.isArray(res) && res.length ? res[0].messages : null;
}

/* Antes de confiar no resultado, prova que a checagem está viva: um trecho com
   um símbolo inexistente TEM que acusar. Se ele passa, o resultado dos arquivos
   de verdade não vale nada. */
function autoteste() {
  const m = lint('var a = 1; console.log(a + simboloQueNaoExiste);');
  return m !== null && m.some((x) => x.ruleId === 'no-undef');
}

const alvos = process.argv.slice(2).filter((a) => a !== '--autoteste');
let falhou = false;

try {
  const vivo = autoteste();
  if (!vivo) {
    console.warn('! ESLint não está checando (faltando? rode: npm i -D eslint)');
    console.warn('  Caindo pro node --check, que só pega erro de sintaxe.');
  }
  if (process.argv.includes('--autoteste')) {
    console.log(vivo ? '✓ checagem viva' : '✗ checagem morta');
    process.exit(vivo ? 0 : 1);
  }
  if (!alvos.length) {
    console.error('uso: node ferramentas/checar-html.mjs arquivo.html');
    process.exit(2);
  }

  for (const arq of alvos) {
    /* Só os <script> sem src — os externos não estão no arquivo pra conferir. */
    const blocos = [...readFileSync(arq, 'utf8')
      .matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)].map((m) => m[1]);
    if (!blocos.length) { console.log(`· ${arq}: sem script embutido`); continue; }
    const js = blocos.join('\n;\n');

    if (vivo) {
      const msgs = lint(js);
      const erros = (msgs ?? []).filter((m) => m.severity === 2);
      if (erros.length) {
        falhou = true;
        console.error(`✗ ${arq}`);
        for (const e of erros) console.error(`  ${e.line}:${e.column}  ${e.message}  ${e.ruleId ?? ''}`);
      } else console.log(`✓ ${arq}`);
      continue;
    }

    writeFileSync(tmpJs, js);
    try { execFileSync(process.execPath, ['--check', tmpJs], { stdio: 'pipe' }); console.log(`✓ ${arq} (só sintaxe)`); }
    catch (e) { falhou = true; console.error(`✗ ${arq}\n` + String(e.stderr ?? '').trim()); }
  }
} finally {
  rmSync(tmpJs, { force: true });
}

process.exit(falhou ? 1 : 0);
