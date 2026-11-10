# Plataforma de Acompanhamento de Obra · SE Campo Formoso II 500 kV

Plataforma web em arquivo único (`SE_Campo_Formoso_II_3D.html`) com Cronograma, Curva S, Placar, Acompanhamento 3D e EAP.
Abra o HTML no navegador (precisa de internet para as bibliotecas da CDN).

O handoff completo está em [`RESUMO_PARA_CLAUDE_CODE.md`](RESUMO_PARA_CLAUDE_CODE.md).

## Comandos
```bash
npm install          # three e xlsx locais, usados só no teste de interface
npm run extract      # regenera items.json a partir do 3D
npm run build-xlsx   # regenera SE_CFII_Avanco_Obra.xlsx (requer openpyxl)
npm run test:ui      # teste no navegador (Playwright); capturas em test-output/
```
