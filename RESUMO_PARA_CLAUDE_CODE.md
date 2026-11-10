# Handoff – Plataforma de Acompanhamento de Obra · SE Campo Formoso II 500 kV (CONSAG)

## 1. Contexto
- **Obra:** ampliação da SE Campo Formoso II 500 kV (Eletrobras/Tractebel, Leilão ANEEL 001/2024, Lote 05), executada pela CONSAG Engenharia.
- **Escopo modelado:** somente as partes hachuradas dos desenhos SECFD.01-500-PB-EM.PTA-0076-00 (planta) e CTE-0078-00 (cortes A-A/B-B):
  - vão de entrada da LT Bom Nome II C1;
  - banco de 3 reatores a ar + 1 reserva;
  - lado da barra (chaves, disjuntores, TCs);
  - casa de controle, GMG1/GMG2 e SA1.
- **Objetivo:** plataforma web de acompanhamento visual da obra, com menu lateral de 5 itens:
  1. Cronograma
  2. Curva S
  3. Placar
  4. Acompanhamento 3D
  5. EAP

## 2. Arquivos do pacote
| Arquivo | O que é |
|---|---|
| `SE_Campo_Formoso_II_3D.html` | A plataforma inteira em 1 arquivo HTML (≈1.530 linhas, sem build). Fonte principal. |
| `SE_CFII_Avanco_Obra.xlsx` | Planilha de avanço do 3D: abas Instruções, Resumo, Quantitativo, Avanço e Equipamentos. |
| `Cronograma_SE_Campo_Formoso_CurvaS.xlsx` | Excel padrão do cronograma, fornecido pelo usuário. Alimenta os itens 1 e 2. |
| `tools/extract.mjs` | Roda o JS do HTML no Node com o three.js "mockado" e exporta a lista de IDs (`items.json`). |
| `tools/build_xlsx.py` | Gera `SE_CFII_Avanco_Obra.xlsx` a partir de `items.json` (openpyxl, com fórmulas). |
| `items.json` | Lista atual: 132 itens com ID e 254 atividades. |

**Regra de ouro:** os IDs nascem no código do 3D. Se mudar o modelo, rode `extract.mjs` e depois `build_xlsx.py`, para a planilha nunca divergir do 3D.
```
node tools/extract.mjs SE_Campo_Formoso_II_3D.html items.json
python3 tools/build_xlsx.py items.json SE_CFII_Avanco_Obra.xlsx
```
Depois disso, recalcule a planilha com o LibreOffice para gravar os valores das fórmulas.

## 3. Stack e estrutura do HTML
- **Arquivo único:** sem framework e sem build. O CSS e o JS (`<script type="module">`) ficam inline. Os logos CONSAG estão embutidos em base64: "full color negativo" na barra lateral; o "positivo" não está mais em uso.
- **Bibliotecas por CDN, com versão fixa:**
  - `three@0.160.0`, via importmap no jsDelivr: `OrbitControls`, `CSS2DRenderer`;
  - `xlsx@0.18.5` (SheetJS), via cdnjs;
  - fonte Montserrat, via Google Fonts.

  Por isso a plataforma precisa de internet. Item pendente: empacotar para uso offline.
- **Roteamento por hash:** `#cronograma`, `#curva-s`, `#placar`, `#acompanhamento-3d`, `#eap`. A função `showPage(k)` alterna os `<main class="view" id="view-…">` e os blocos de ações do cabeçalho (`header.top [data-for]`). O render loop do 3D pausa fora da view 3D.
- **Persistência:** só `localStorage`, por navegador:
  - `cfii-prog`: % por atividade do 3D;
  - `cfii-cron`: cronograma parseado + data de corte.

  Não há backend.
- **Identidade visual (manual CONSAG):**
  - fonte Montserrat;
  - azul `#3074F6`, marinho `#06172B`, `#283344`, Pantone 303 `#001434` (fundo do 3D), ciano `#00A1D4` (seleção e data de corte), `#2D758E`;
  - status: verde claro `#C6EFCE`/`#9EE0AB`, amarelo claro `#FDF0A6`/`#F6E27A`, cinza `#E6EBF1`.

  Botões principais em marinho com texto branco, porque o azul com texto branco pequeno não passa no contraste.

## 4. Item 4 – Acompanhamento 3D
- **Coordenadas (m):**
  - X = afastamento a partir do alambrado da LT;
  - Z = a partir do eixo 4 (eixo 5 = 28 m);
  - Y = altura.
  - Fases: VM z = 5,5, BR z = 14, AZ z = 22,5. Reatores em z = 8,5 / 20 / 31,5 / 43 (reserva), com X = 61,5.
- **Construtores de equipamento:** `PR`, `TPC`, `IP`, `BB`, `CHV`, `TC`, `CHS`, `DJ`, `REATOR`, `IPN`, `TCN`, `gantry`, `fence`, casa/GMG/SA. Alguns seguem modelos do fornecedor:
  - reator a ar;
  - chave seccionadora de abertura vertical, representada **fechada**;
  - bobina de bloqueio com sapata, bloco, coluna, placa octogonal e 6 isoladores;
  - TC tipo cabeça com terminais P1/P2.
- **Bobinas de bloqueio:** ficam nas fases BR e AZ, em X = 15. A fase VM tem isolador de pedestal nesse ponto.
- **Condutores 4×CAA 954 MCM "Rail":** função `bundle()`, com 4 subcondutores em quadrado de 457 mm e espaçador a cada ~6 m.
  - Conector H nos equipamentos, via `hconn`.
  - TCs e chaves seccionadoras recebem o feixe direto no quadro terminal de 4 grampos (`termFrame`), usando pontos `RAW()` sem elevação.
- **Etapas por componente:** dentro de cada construtor, `mark(g,'F'|'S'|'E')` separa os filhos em Fundação, Suporte e Equipamento. Cabos usam a etapa C.
- **Registro:** `E(obj, code, area, ph, name, extra, layer)`.
  - ID = `ÁREA-FASE-TIPO-NN`, por exemplo `LT-VM-TC-01`.
  - Atividade = ID + `.F`/`.S`/`.E`/`.C`.
  - Áreas: LT (entrada de linha), RE (reatores), BA (vão da barra), SX (serviços auxiliares).
  - Fases: VM, BR, AZ, RS (reserva), 3F (comum).
  - A tabela `TIPOS` define a descrição e as etapas aplicáveis de cada código.
- **Modo "Avanço da obra":**
  - concluído (100 %) = verde claro;
  - 1–99 % = amarelo claro;
  - 0 % = fantasma.

  Pingados a retirar e conexão da reserva (Nota 5) não são rastreados e só aparecem no "Modelo completo".
- **Painel do 3D:**
  - KPIs por etapa;
  - busca por ID;
  - filtros de etapa e fase;
  - camadas;
  - "Atividades em andamento";
  - "Quantitativo do escopo": Prev. / Real. (=100 %) / Saldo / Avanço (média do %) por tipo × etapa.
- **Leitura da planilha:** aba "Avanço", colunas "ID Atividade" e "% Avanço". Um número ≤ 1 é lido como fração.

## 5. Item 1 – Cronograma
Núcleo puro e testável: `CRC.parse(book)` + `CRC.compute(P, cut)`. A interface fica em `CR` / `crRender` / `crGantt`.
- **Abas lidas:**
  - **Cronograma:** Disciplina, Grupo, Atividade, Início/Término Linha Base, Início, Término, Trabalho, Peso. Só entram linhas com Disciplina preenchida (48 atividades). Os pesos são normalizados para somar 1.
  - **Feriados:** dias com "Considerar? = S".
  - **Distribuição Real.:** a coluna A "Linha" é o número da linha no Cronograma. A linha 3 traz as segundas e a linha 4 os domingos; as semanas começam na coluna F.
- **Calendário:** segunda a sábado, menos feriados.
- **Cálculos:**
  - `% Prev.` = dias úteis da LB até a data de corte ÷ dias úteis totais da LB.
  - `% Real.` = soma dos pesos semanais com segunda ≤ data de corte.
- **Data de corte:** por padrão, o domingo da última semana com realizado lançado. É editável e compartilhada com o item 2.
- **Tela:**
  - KPIs;
  - resumo por disciplina (Peso, Avanço previsto, Avanço real, Desvio; verde se ≥ 0, amarelo até −5 p.p., vermelho abaixo);
  - Gantt por Disciplina → Grupo, que se abre em atividades, com colunas Peso, % Prev., % Real., Início/Término LB, Início/Término.
- **Gantt:**
  - barras: LB em cinza, tendência com contorno azul, avanço preenchido em azul × % Real;
  - escala mensal; clicar no mês abre em semanas (S01…), e o botão "Semanal" abre todos;
  - colunas da esquerda fixas e rolagem automática até a data de corte.
- **Validação com o arquivo, em 16/01/2028:** previsto 19,68 % (= aba "Curva S", semana 9) e real 19,25 % (= linha "Conferência" da Distribuição Real.).

## 6. Item 2 – Curva S
- **Abas:** "Distribuição LB", "Distribuição Tendência" e "Distribuição Real.", com o mesmo layout. Em `parse`, a função `readDist()` alinha as semanas pela segunda-feira e guarda o resultado em `P.dist`.
- **Linhas descartadas:** linhas sem "Linha", como "Total semanal" e "Conferência".
- **Recorte:** filtros de Disciplina e Grupo; o acumulado é normalizado pelo peso do recorte.
- **Realizado:** vai até a última semana com lançamento.
- **KPIs:** LB, Real, Tendência, Desvio, IDP = Real ÷ LB, semana atual e término LB × tendência.
- **Gráfico SVG:**
  - 3 curvas;
  - barras do período;
  - data de corte;
  - tooltip no hover;
  - escala semanal ou mensal (no mensal, agrupa as semanas pelo mês do domingo).
- **Tabela:** semanal ou mensal.
- **Validação:** LB acumulado na semana 9 = 19,68 %, Real = 19,25 %, LB final = 100 %, Tendência final ≈ 99,98 %.
- **Ponto aberto:** o término na Curva S (24/12/28, domingo da última semana) difere do Cronograma (22/12/28, última atividade). Falta decidir qual critério usar nos dois lugares.
- **Compatibilidade:** quem carregou o cronograma antes desta versão precisa recarregar o Excel, porque `P.dist` não existia.

## 7. Itens 3 e 5 (ainda como placeholder no HTML; layout aprovado no mockup)
- **3. Placar:** cartões por equipamento principal (montados / previstos, com fundação e suporte), vindos da planilha de avanço do 3D (`PROG`).
- **5. EAP:** árvore obra → área → etapa → atividade (ID.etapa), com contagens reais. Por área:

| Área | IDs | Atividades | Fundação | Suporte | Montagem | Cabos |
|---|---|---|---|---|---|---|
| LT | 52 | 91 | 20 | 21 | 25 | 25 |
| RE | 34 | 82 | 24 | 25 | 24 | 9 |
| BA | 42 | 72 | 15 | 15 | 24 | 18 |
| SX | 4 | 9 | 4 | 1 | 4 | – |

- **Mockup:** fica num canvas Design no claude.ai (artboards "Plataforma com navegação lateral (v2)" e "2. Curva S"). O link é privado, então peça ao usuário prints ou acesso.

## 8. Pendências e próximos passos sugeridos
1. Implementar o **Placar** e a **EAP** funcionais.
2. Tela de **atualização em campo** (celular, leitura de QR por ID). Exige armazenamento compartilhado, porque hoje é só `localStorage` + Excel.
3. **Data de corte = hoje** quando a obra começar (hoje é a última semana com realizado, porque o cronograma começa em nov/2027).
4. Unificar o **critério de término** entre Cronograma e Curva S.
5. **Pesos** no avanço do 3D (hoje é média simples das atividades).
6. Pacote **offline**, com three.js e SheetJS embutidos.
7. Trocar as medidas estimadas pelos **dimensionais dos fornecedores**: reator, BB, TC, chaves. Hoje vêm de proporção de imagem e da escala 1:300 dos cortes.

## 9. Premissas do modelo 3D
- Distâncias dos cortes A-A/B-B; alturas não cotadas lidas na escala 1:300. Viga dos pórticos ≈ 23–25 m, mastro ≈ 34 m, barramento baixo 8–10 m.
- Fase AZ assumida igual à VM, porque não há corte da AZ.
- Casa de controle, GMGs e SA1 posicionados pela planta, sem cotas de amarração. As 4 "baias" ao lado da casa têm finalidade a confirmar.
- O pórtico do eixo A (X = 96,5) e o pórtico em X = 163,5 são existentes: aparecem como fantasma, fora do escopo.

## 10. Como testar sem internet
- **Lógica:** `node tools/extract.mjs …` executa todo o JS com um Proxy no lugar do three.js e valida que o modelo monta sem erro.
- **Interface:** com Playwright/Chromium, remova o importmap e o script do SheetJS, injete o mesmo Proxy e pré-carregue `localStorage['cfii-cron']` com o resultado de `CRC.parse` gerado no Node a partir do Excel (convertido para arrays, com datas como serial do Excel).
