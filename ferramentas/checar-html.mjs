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
import { join, dirname } from 'node:path';
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

/* Classes que existem só como gancho de JS ou espaçador de grade, e por isso
   nunca vão ter regra própria. Lista curta e explícita: o dia que uma delas
   ganhar CSS, some daqui. */
const SEM_CSS_DE_PROPOSITO = new Set([
  'mp-canto',                             // célula vazia da grade do mapa
  'soFunil', 'soFonte', 'soCriativos',    // visibilidade por JS, sobre .grp
]);

function classesOrfas(html, arq) {
  /* O CSS pode estar em <style> ou num arquivo linkado — o quiz usa
     quiz-warm.css, e olhar só o inline reportaria as 200 classes dele como
     órfãs. Um alarme que dispara 200 vezes é um alarme desligado. */
  let css = [...html.matchAll(/<style[^>]*>([\s\S]*?)<\/style>/g)].map((m) => m[1]).join('\n');
  for (const m of html.matchAll(/<link[^>]*rel=["']stylesheet["'][^>]*>/g)) {
    const href = (m[0].match(/href=["']([^"']+)["']/) || [])[1];
    if (!href || /^https?:|^\/\//.test(href)) continue;   // externo: não dá pra ler
    try { css += '\n' + readFileSync(join(dirname(arq), href.split('?')[0]), 'utf8'); }
    catch { return []; }   /* folha que não abre = não sei dizer; melhor calar */
  }
  const resto = html.replace(/<style[^>]*>[\s\S]*?<\/style>/g, '');
  if (!css.trim()) return [];

  const usadas = new Set();
  for (const m of resto.matchAll(/class="([^"]*)"/g)) {
    /* Só o pedaço LITERAL, antes da primeira concatenação: em
       class="' + cls + '" o que vem depois é nome de variável, não de classe,
       e reportá-lo seria alarme falso toda vez. */
    const literal = m[1].split(/['"`+${]/)[0];
    const toks = literal.split(/\s+/).filter(Boolean);
    /* Se a concatenação começa colada no último token, ele está cortado no meio
       — 's1-v' + TELA é a classe s1-v3, não uma classe chamada s1-v. */
    if (literal !== m[1] && !/\s$/.test(literal)) toks.pop();
    for (const t of toks) if (/^[a-zA-Z][\w-]*$/.test(t)) usadas.add(t);
  }
  /* O [,)] no fim exige que a string seja o argumento INTEIRO: em
     classList.add('s1-v' + TELA) o nome real é s1-v3, e pegar só o pedaço
     literal reportaria uma classe que nunca existiu. */
  for (const m of resto.matchAll(/classList\.(?:add|remove|toggle)\('([\w-]+)'\s*[,)]/g)) usadas.add(m[1]);

  const definidas = new Set([...css.matchAll(/\.([a-zA-Z][\w-]*)/g)].map((m) => m[1]));
  return [...usadas].filter((u) => !definidas.has(u) && !SEM_CSS_DE_PROPOSITO.has(u)).sort();
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
    const html = readFileSync(arq, 'utf8');

    /* Classe usada sem nenhuma regra de CSS. Mesmo estrago do no-undef e
       invisível do mesmo jeito: `.mp-grade` sumiu num commit e o mapa de calor
       virou uma lista de números empilhados, sem erro nenhum no console.
       Aconteceu com `.livre` e `.fcheck` no mesmo commit, pela mesma razão —
       recortar uma região do CSS por âncoras leva os vizinhos junto. */
    for (const orfa of classesOrfas(html, arq)) {
      falhou = true;
      console.error(`✗ ${arq}\n  classe "${orfa}" usada no HTML mas sem regra de CSS`);
    }

    /* Só os <script> sem src — os externos não estão no arquivo pra conferir. */
    const blocos = [...html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)]
      .map((m) => m[1]);
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
