---
title: Referência da API de scripts do Office
description: Uma visão geral das APIs javaScript Office scripts.
ms.date: 06/29/2020
ms.openlocfilehash: 3ce3344fb49b2811719feb13f8fb4118f1a20060db9d85a06d1be939f22bf3c5
ms.sourcegitcommit: 6a0182075a558c4fe664fedfee08fea76513b192
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/11/2021
ms.locfileid: "58183423"
---
# <a name="office-scripts-api-reference"></a>Referência da API de scripts do Office

A OFFICE scripts permite automatizar tarefas comuns em Excel na Web. Use esta documentação de referência para saber mais sobre as classes, métodos e outros tipos disponíveis para seus scripts. Todos os objetos acessíveis Office scripts podem ser encontrados no índice de conteúdo à esquerda da página.

> [!NOTE]
> Se você estiver procurando as APIs JavaScript para o desenvolvimento de Office de Office, visite Office referência da API JavaScript de [complementos.](/javascript/api/overview?view=excel-js-preview)

## <a name="common-classes"></a>Classes comuns

A lista a seguir divide os conceitos básicos do modelo de objeto Office Scripts. Isso mostra as classes comuns e como elas se relacionam entre si.

- Uma [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) contém uma ou mais [Planilhas](/javascript/api/office-scripts/excelscript/excelscript.worksheet).
- Uma [Planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet) concede acesso a células por meio de objetos de [Intervalo](/javascript/api/office-scripts/excelscript/excelscript.range).
- Um [Intervalo](/javascript/api/office-scripts/excelscript/excelscript.range) representa um grupo de células contíguas.
- Os [Intervalos](/javascript/api/office-scripts/excelscript/excelscript.range) são usados para criar e colocar [Tabelas](/javascript/api/office-scripts/excelscript/excelscript.table), [Gráficos](/javascript/api/office-scripts/excelscript/excelscript.chart), [Formas](/javascript/api/office-scripts/excelscript/excelscript.shape) e outras visualizações de dados ou objetos da organização.
- Uma [Planilha contém](/javascript/api/office-scripts/excelscript/excelscript.worksheet) matrizes preenchidas com os objetos presentes na planilha individual.
- Uma [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) contém matrizes de alguns desses objetos de dados para toda a Workbook.

Para obter mais informações sobre o modelo de objeto Office Scripts, visite [Scripting fundamentals for Office Scripts no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>Confira também

- [Sobre Office scripts](/office/dev/scripts/overview/excel)
- [Gravar, editar e criar scripts do Office no Excel na Web](/office/dev/scripts/tutorials/excel-tutorial)
- [Fundamentos de script para scripts do Office no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)
