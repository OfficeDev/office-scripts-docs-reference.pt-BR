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
# <a name="office-scripts-api-reference"></a><span data-ttu-id="61690-103">Referência da API de scripts do Office</span><span class="sxs-lookup"><span data-stu-id="61690-103">Office Scripts API reference</span></span>

<span data-ttu-id="61690-104">A API de scripts do Office permite automatizar tarefas comuns no Excel na Web.</span><span class="sxs-lookup"><span data-stu-id="61690-104">The Office Scripts API lets you automate common tasks in Excel on the web.</span></span> <span data-ttu-id="61690-105">Use esta documentação de referência para saber mais sobre as classes, métodos e outros tipos disponíveis para seus scripts.</span><span class="sxs-lookup"><span data-stu-id="61690-105">Use this reference documentation to learn more about the classes, methods, and other types available for your scripts.</span></span> <span data-ttu-id="61690-106">Todos os objetos acessíveis através de scripts do Office podem ser encontrados no sumário à esquerda da página.</span><span class="sxs-lookup"><span data-stu-id="61690-106">All the objects accessible through Office Scripts can be found in the table of contents on the left of the page.</span></span>

> [!NOTE]
> <span data-ttu-id="61690-107">Se você estiver procurando as APIs JavaScript para desenvolver suplementos do Office, visite a referência da [API JavaScript de suplementos do Office](/javascript/api/overview?view=excel-js-preview).</span><span class="sxs-lookup"><span data-stu-id="61690-107">If you're looking for the JavaScript APIs for developing Office Add-ins, visit the [Office Add-ins JavaScript API reference](/javascript/api/overview?view=excel-js-preview).</span></span>

## <a name="common-classes"></a><span data-ttu-id="61690-108">Classes comuns</span><span class="sxs-lookup"><span data-stu-id="61690-108">Common classes</span></span>

<span data-ttu-id="61690-109">A lista a seguir divide as noções básicas do modelo de objeto de scripts do Office.</span><span class="sxs-lookup"><span data-stu-id="61690-109">The following list breaks down the basics of the Office Scripts object model.</span></span> <span data-ttu-id="61690-110">Isso mostra as classes comuns e como elas se relacionam entre si.</span><span class="sxs-lookup"><span data-stu-id="61690-110">This shows the common classes and how they relate to one another.</span></span>

- <span data-ttu-id="61690-111">Uma [Pasta de trabalho](/javascript/api/office-scripts/excel/excelscript.workbook) contém uma ou mais [Planilhas](/javascript/api/office-scripts/excel/excelscript.worksheet).</span><span class="sxs-lookup"><span data-stu-id="61690-111">A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains one or more [Worksheets](/javascript/api/office-scripts/excel/excelscript.worksheet).</span></span>
- <span data-ttu-id="61690-112">Uma [Planilha](/javascript/api/office-scripts/excel/excelscript.worksheet) concede acesso a células por meio de objetos de [Intervalo](/javascript/api/office-scripts/excel/excelscript.range).</span><span class="sxs-lookup"><span data-stu-id="61690-112">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) gives access to cells through [Range](/javascript/api/office-scripts/excel/excelscript.range) objects.</span></span>
- <span data-ttu-id="61690-113">Um [Intervalo](/javascript/api/office-scripts/excel/excelscript.range) representa um grupo de células contíguas.</span><span class="sxs-lookup"><span data-stu-id="61690-113">A [Range](/javascript/api/office-scripts/excel/excelscript.range) represents a group of contiguous cells.</span></span>
- <span data-ttu-id="61690-114">Os [Intervalos](/javascript/api/office-scripts/excel/excelscript.range) são usados para criar e colocar [Tabelas](/javascript/api/office-scripts/excel/excelscript.table), [Gráficos](/javascript/api/office-scripts/excel/excelscript.chart), [Formas](/javascript/api/office-scripts/excel/excelscript.shape) e outras visualizações de dados ou objetos da organização.</span><span class="sxs-lookup"><span data-stu-id="61690-114">[Ranges](/javascript/api/office-scripts/excel/excelscript.range) are used to create and place [Tables](/javascript/api/office-scripts/excel/excelscript.table), [Charts](/javascript/api/office-scripts/excel/excelscript.chart), [Shapes](/javascript/api/office-scripts/excel/excelscript.shape), and other data visualization or organization objects.</span></span>
- <span data-ttu-id="61690-115">Uma [planilha](/javascript/api/office-scripts/excel/excelscript.worksheet) contém matrizes preenchidas com os objetos que estão presentes na planilha individual.</span><span class="sxs-lookup"><span data-stu-id="61690-115">A [Worksheet](/javascript/api/office-scripts/excel/excelscript.worksheet) contains arrays filled with those objects that are present in the individual sheet.</span></span>
- <span data-ttu-id="61690-116">Uma [pasta de trabalho](/javascript/api/office-scripts/excel/excelscript.workbook) contém matrizes de alguns desses objetos de dados para a pasta de trabalho inteira.</span><span class="sxs-lookup"><span data-stu-id="61690-116">A [Workbook](/javascript/api/office-scripts/excel/excelscript.workbook) contains arrays of some of those data objects for the entire Workbook.</span></span>

<span data-ttu-id="61690-117">Para obter mais informações sobre o modelo de objeto de scripts do Office, visite [os conceitos básicos de script para scripts do Office no Excel na Web](/office/dev/scripts/develop/scripting-fundamentals)</span><span class="sxs-lookup"><span data-stu-id="61690-117">For more information about the Office Scripts object model, visit [Scripting fundamentals for Office Scripts in Excel on the web](/office/dev/scripts/develop/scripting-fundamentals)</span></span>

## <a name="see-also"></a><span data-ttu-id="61690-118">Confira também</span><span class="sxs-lookup"><span data-stu-id="61690-118">See also</span></span>

- [<span data-ttu-id="61690-119">Sobre scripts do Office</span><span class="sxs-lookup"><span data-stu-id="61690-119">About Office Scripts</span></span>](/office/dev/scripts/overview/excel)
- [<span data-ttu-id="61690-120">Gravar, editar e criar scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="61690-120">Record, edit, and create Office Scripts in Excel on the web</span></span>](/office/dev/scripts/tutorials/excel-tutorial)
- [<span data-ttu-id="61690-121">Fundamentos de script para scripts do Office no Excel na Web</span><span class="sxs-lookup"><span data-stu-id="61690-121">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](/office/dev/scripts/develop/scripting-fundamentals)
