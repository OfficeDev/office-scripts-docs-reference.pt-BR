---
title: Referência da API de scripts do Office
description: Uma visão geral das APIs JavaScript de scripts assíncronos do Office.
ms.date: 06/29/2020
ms.openlocfilehash: 3d8d37b30d9535e8b6a56a08c44f9034cb599f31
ms.sourcegitcommit: 9c4c4c213a203e58c55eb3d84d7d92fa527f3eb8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/01/2020
ms.locfileid: "45003951"
---
# <a name="office-scripts-async-api-reference"></a>Referência da API assíncrona de scripts do Office

A API Async de scripts do Office suporta scripts mais antigos feitos durante a fase de visualização de scripts do Office. Use esta documentação de referência para saber mais sobre as classes, os métodos e outros tipos usados por esses scripts mais antigos. Todos os objetos acessíveis através de scripts do Office podem ser encontrados no sumário à esquerda da página.

> [!IMPORTANT]
> É altamente recomendável criar novos scripts com as APIs de scripts padrão do Office. Se você estiver fazendo ou editando um novo script, alterne para a [versão não assíncrona](?view=office-scripts) das APIs.

## <a name="common-classes"></a>Classes comuns

A lista a seguir divide as noções básicas do modelo de objeto de scripts do Office. Isso mostra as classes comuns e como elas se relacionam entre si.

- Uma [pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) contém uma ou mais [planilhas](/javascript/api/office-scripts/excelscript/excelscript.worksheet) em uma [worksheetcollection](/javascript/api/office-scripts/excelscript/excelscript.worksheetcollection).
- Uma [Planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet) concede acesso a células por meio de objetos de [Intervalo](/javascript/api/office-scripts/excelscript/excelscript.range).
- Um [Intervalo](/javascript/api/office-scripts/excelscript/excelscript.range) representa um grupo de células contíguas.
- Os [Intervalos](/javascript/api/office-scripts/excelscript/excelscript.range) são usados para criar e colocar [Tabelas](/javascript/api/office-scripts/excelscript/excelscript.table), [Gráficos](/javascript/api/office-scripts/excelscript/excelscript.chart), [Formas](/javascript/api/office-scripts/excelscript/excelscript.shape) e outras visualizações de dados ou objetos da organização.
- Uma [planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contém coleções desses objetos de dados (como um [chartcollection](/javascript/api/office-scripts/excelscript/excelscript.chartcollection)) que estão presentes na planilha individual.
- As [pastas de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) contêm coleções de alguns desses objetos de dados (como [TableCollection](/javascript/api/office-scripts/excelscript/excelscript.tablecollection)) para a pasta de [trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook)inteira.

Para obter mais informações sobre o modelo de objeto de scripts do Office, visite [os conceitos básicos de script para scripts do Office no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>Também consulte

- [Usando as APIs assíncronas de scripts do Office para suportar scripts herdados](/office/dev/scripts/develop/excel-async-model)
- [Sobre scripts do Office](/office/dev/scripts/overview/excel)
- [Gravar, editar e criar scripts do Office no Excel na Web](/office/dev/scripts/tutorials/excel-tutorial)
- [Fundamentos de script para scripts do Office no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)
