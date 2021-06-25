# <a name="office-scripts-api-documentation-tools"></a>Office Ferramentas de documentação da API de scripts

Essas ferramentas ajudam a dar suporte à documentação Office SCripts e à equipe por trás dela. Siga estas instruções para executar as ferramentas nesta pasta.

## <a name="coverage-tester"></a>coverage-tester

Esta ferramenta fornece uma visão geral da cobertura de documentação para cada API. Cada API é avaliada para a qualidade da documentação e a presença de código de exemplo. As métricas de qualidade ainda estão em desenvolvimento.

A saída dessa ferramenta é um `.csv` arquivo.

### <a name="coverage-tester-instructions"></a>Instruções do testador de cobertura

1. Clonar ou bifurcar o repo.
1. Em uma janela de comando, vá para `/office-scripts-docs-reference/generate-docs/tools`
1. Executar `npm install`
1. Executar `npm run build`
1. Executar `node coverage-tester`
1. Abra "Api Coverage Report.csv"
