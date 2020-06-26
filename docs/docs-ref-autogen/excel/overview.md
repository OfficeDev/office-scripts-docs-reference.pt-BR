---
title: Referência da API de scripts do Office
description: Uma visão geral das APIs JavaScript de scripts do Office.
ms.date: 06/17/2020
ms.openlocfilehash: 5634d0e5f68464655054ad1c09eb7931e0da62d4
ms.sourcegitcommit: 163b26a43411ad7f13a01237efe9b8d6de656b47
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44879783"
---
# <a name="office-scripts-api-reference"></a>Referência da API de scripts do Office

A API de scripts do Office permite automatizar tarefas comuns no Excel na Web. Use esta documentação de referência para saber mais sobre as classes, métodos e outros tipos disponíveis para seus scripts. Todos os objetos acessíveis através de scripts do Office podem ser encontrados no sumário à esquerda da página.

> [!NOTE]
> Se você estiver procurando as APIs JavaScript para desenvolver suplementos do Office, visite a referência da [API JavaScript de suplementos do Office](/javascript/api/overview?view=excel-js-preview).

## <a name="common-classes"></a>Classes comuns

A lista a seguir divide as noções básicas do modelo de objeto de scripts do Office. Isso mostra as classes comuns e como elas se relacionam entre si.

- Uma [Pasta de trabalho](/javascript/api/office-scripts/excel/excelscript.workbook) contém uma ou mais [Planilhas](/javascript/api/office-scripts/excel/excelscript.worksheet).
- Uma [Planilha](/javascript/api/office-scripts/excel/excelscript.worksheet) concede acesso a células por meio de objetos de [Intervalo](/javascript/api/office-scripts/excel/excelscript.range).
- Um [Intervalo](/javascript/api/office-scripts/excel/excelscript.range) representa um grupo de células contíguas.
- Os [Intervalos](/javascript/api/office-scripts/excel/excelscript.range) são usados para criar e colocar [Tabelas](/javascript/api/office-scripts/excel/excelscript.table), [Gráficos](/javascript/api/office-scripts/excel/excelscript.chart), [Formas](/javascript/api/office-scripts/excel/excelscript.shape) e outras visualizações de dados ou objetos da organização.
- Uma [planilha](/javascript/api/office-scripts/excel/excelscript.worksheet) contém matrizes preenchidas com os objetos que estão presentes na planilha individual.
- Uma [pasta de trabalho](/javascript/api/office-scripts/excel/excelscript.workbook) contém matrizes de alguns desses objetos de dados para a pasta de trabalho inteira.

Para obter mais informações sobre o modelo de objeto de scripts do Office, visite [os conceitos básicos de script para scripts do Office no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>Confira também

- [Sobre scripts do Office](/office/dev/scripts/overview/excel)
- [Gravar, editar e criar scripts do Office no Excel na Web](/office/dev/scripts/tutorials/excel-tutorial)
- [Fundamentos de script para scripts do Office no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)
