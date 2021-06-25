# <a name="office-scripts-api-documentation-tools"></a><span data-ttu-id="209dd-101">Office Ferramentas de documentação da API de scripts</span><span class="sxs-lookup"><span data-stu-id="209dd-101">Office Scripts API Documentation Tools</span></span>

<span data-ttu-id="209dd-102">Essas ferramentas ajudam a dar suporte à documentação Office SCripts e à equipe por trás dela.</span><span class="sxs-lookup"><span data-stu-id="209dd-102">These tools help support the Office SCripts documentation and the team behind it.</span></span> <span data-ttu-id="209dd-103">Siga estas instruções para executar as ferramentas nesta pasta.</span><span class="sxs-lookup"><span data-stu-id="209dd-103">Follow these instructions to run the tools in this folder.</span></span>

## <a name="coverage-tester"></a><span data-ttu-id="209dd-104">coverage-tester</span><span class="sxs-lookup"><span data-stu-id="209dd-104">coverage-tester</span></span>

<span data-ttu-id="209dd-105">Esta ferramenta fornece uma visão geral da cobertura de documentação para cada API.</span><span class="sxs-lookup"><span data-stu-id="209dd-105">This tool gives an overview of the documentation coverage for each API.</span></span> <span data-ttu-id="209dd-106">Cada API é avaliada para a qualidade da documentação e a presença de código de exemplo.</span><span class="sxs-lookup"><span data-stu-id="209dd-106">Each API is assessed for documentation quality and the presence of sample code.</span></span> <span data-ttu-id="209dd-107">As métricas de qualidade ainda estão em desenvolvimento.</span><span class="sxs-lookup"><span data-stu-id="209dd-107">The quality metrics are still in development.</span></span>

<span data-ttu-id="209dd-108">A saída dessa ferramenta é um `.csv` arquivo.</span><span class="sxs-lookup"><span data-stu-id="209dd-108">The output of this tool is a `.csv` file.</span></span>

### <a name="coverage-tester-instructions"></a><span data-ttu-id="209dd-109">Instruções do testador de cobertura</span><span class="sxs-lookup"><span data-stu-id="209dd-109">coverage-tester Instructions</span></span>

1. <span data-ttu-id="209dd-110">Clonar ou bifurcar o repo.</span><span class="sxs-lookup"><span data-stu-id="209dd-110">Clone or fork the repo.</span></span>
1. <span data-ttu-id="209dd-111">Em uma janela de comando, vá para `/office-scripts-docs-reference/generate-docs/tools`</span><span class="sxs-lookup"><span data-stu-id="209dd-111">In a command window, go to `/office-scripts-docs-reference/generate-docs/tools`</span></span>
1. <span data-ttu-id="209dd-112">Executar `npm install`</span><span class="sxs-lookup"><span data-stu-id="209dd-112">Run `npm install`</span></span>
1. <span data-ttu-id="209dd-113">Executar `npm run build`</span><span class="sxs-lookup"><span data-stu-id="209dd-113">Run `npm run build`</span></span>
1. <span data-ttu-id="209dd-114">Executar `node coverage-tester`</span><span class="sxs-lookup"><span data-stu-id="209dd-114">Run `node coverage-tester`</span></span>
1. <span data-ttu-id="209dd-115">Abra "Api Coverage Report.csv"</span><span class="sxs-lookup"><span data-stu-id="209dd-115">Open “API Coverage Report.csv”</span></span>
