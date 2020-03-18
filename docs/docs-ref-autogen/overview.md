---
title: Referência da API de scripts do Office
description: Uma visão geral das APIs JavaScript de scripts do Office
ms.date: 12/12/2019
ms.openlocfilehash: b3e75cbecac13f41c83f019c681b0682571ead41
ms.sourcegitcommit: 2834ee43a14a4b0953d5fb22749c115a42960e20
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/10/2020
ms.locfileid: "42701515"
---
# <a name="office-scripts-api-reference"></a>Referência da API de scripts do Office

A API de scripts do Office permite automatizar tarefas comuns no Excel na Web. Use esta documentação de referência para saber mais sobre as classes, métodos e outros tipos disponíveis para seus scripts. Todos os objetos acessíveis através de scripts do Office podem ser encontrados no sumário à esquerda da página.

## <a name="common-classes"></a>Classes comuns

A lista a seguir divide as noções básicas do modelo de objeto de scripts do Office. Isso mostra as classes comuns e como elas se relacionam entre si.

- Uma [pasta de trabalho](/javascript/api/office-scripts/excel/excel.workbook) contém uma ou mais [planilhas](/javascript/api/office-scripts/excel/excel.worksheet) em uma [worksheetcollection](/javascript/api/office-scripts/excel/excel.worksheetcollection).
- Uma [planilha](/javascript/api/office-scripts/excel/excel.worksheet) dá acesso a células por meio de objetos [Range](/javascript/api/office-scripts/excel/excel.range) .
- Um [intervalo](/javascript/api/office-scripts/excel/excel.range) representa um grupo de células contíguas.
- Os [intervalos](/javascript/api/office-scripts/excel/excel.range) são usados para criar e colocar [tabelas](/javascript/api/office-scripts/excel/excel.table), [gráficos](/javascript/api/office-scripts/excel/excel.chart), [formas](/javascript/api/office-scripts/excel/excel.shape)e outros objetos de organização ou de visualização de dados.
- Uma [planilha](/javascript/api/office-scripts/excel/excel.worksheet) contém coleções desses objetos de dados (como um [chartcollection](/javascript/api/office-scripts/excel/excel.chartcollection)) que estão presentes na planilha individual.
- [Pastas de trabalho](/javascript/api/office-scripts/excel/excel.workbook). contenha coleções de alguns desses objetos de dados (como [TableCollection](/javascript/api/office-scripts/excel/excel.tablecollection)) para a pasta de [trabalho](/javascript/api/office-scripts/excel/excel.workbook)inteira.

Para obter mais informações sobre o modelo de objeto de scripts do Office, visite [os conceitos básicos de script para scripts do Office no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>Confira também

- [Sobre scripts do Office](/office/dev/scripts/overview/excel)
- [Gravar, editar e criar scripts do Office no Excel na Web](/office/dev/scripts/tutorials/excel-tutorial)
- [Conceitos básicos de script para scripts do Office no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)
