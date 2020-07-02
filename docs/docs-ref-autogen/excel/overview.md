---
title: Referência da API de scripts do Office
description: Uma visão geral das APIs JavaScript de scripts do Office.
ms.date: 06/29/2020
ms.openlocfilehash: 7c4fe97ca35cfb442ebbf9db2e0b03b389185ae8
ms.sourcegitcommit: 9c4c4c213a203e58c55eb3d84d7d92fa527f3eb8
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/01/2020
ms.locfileid: "45004728"
---
# <a name="office-scripts-api-reference"></a>Referência da API de scripts do Office

A API de scripts do Office permite automatizar tarefas comuns no Excel na Web. Use esta documentação de referência para saber mais sobre as classes, métodos e outros tipos disponíveis para seus scripts. Todos os objetos acessíveis através de scripts do Office podem ser encontrados no sumário à esquerda da página.

> [!NOTE]
> Se você estiver procurando as APIs JavaScript para desenvolver suplementos do Office, visite a referência da [API JavaScript de suplementos do Office](/javascript/api/overview?view=excel-js-preview).

## <a name="common-classes"></a>Classes comuns

A lista a seguir divide as noções básicas do modelo de objeto de scripts do Office. Isso mostra as classes comuns e como elas se relacionam entre si.

- Uma [Pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) contém uma ou mais [Planilhas](/javascript/api/office-scripts/excelscript/excelscript.worksheet).
- Uma [Planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet) concede acesso a células por meio de objetos de [Intervalo](/javascript/api/office-scripts/excelscript/excelscript.range).
- Um [Intervalo](/javascript/api/office-scripts/excelscript/excelscript.range) representa um grupo de células contíguas.
- Os [Intervalos](/javascript/api/office-scripts/excelscript/excelscript.range) são usados para criar e colocar [Tabelas](/javascript/api/office-scripts/excelscript/excelscript.table), [Gráficos](/javascript/api/office-scripts/excelscript/excelscript.chart), [Formas](/javascript/api/office-scripts/excelscript/excelscript.shape) e outras visualizações de dados ou objetos da organização.
- Uma [planilha](/javascript/api/office-scripts/excelscript/excelscript.worksheet) contém matrizes preenchidas com os objetos que estão presentes na planilha individual.
- Uma [pasta de trabalho](/javascript/api/office-scripts/excelscript/excelscript.workbook) contém matrizes de alguns desses objetos de dados para a pasta de trabalho inteira.

Para obter mais informações sobre o modelo de objeto de scripts do Office, visite [os conceitos básicos de script para scripts do Office no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)

## <a name="see-also"></a>Também consulte

- [Sobre scripts do Office](/office/dev/scripts/overview/excel)
- [Gravar, editar e criar scripts do Office no Excel na Web](/office/dev/scripts/tutorials/excel-tutorial)
- [Fundamentos de script para scripts do Office no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)
