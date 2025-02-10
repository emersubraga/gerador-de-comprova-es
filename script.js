document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("form").addEventListener("submit", async (event) => {
        event.preventDefault();

        const empresa = document.getElementById("empresa").value.toUpperCase();
        const nota = document.getElementById("nota").value.toUpperCase();
        const ordem = document.getElementById("ordem").value.toUpperCase();
        const data = document.getElementById("data").value.toUpperCase();
        const modeloFile = document.getElementById("modelo").files[0];

        if (!modeloFile) {
            alert("Por favor, selecione um modelo de documento.");
            return;
        }

        if (typeof window.Mammoth === "undefined") {
            console.error("Mammoth.js não foi carregado corretamente.");
            alert("Erro ao carregar Mammoth.js. Verifique sua conexão ou tente novamente.");
            return;
        }

        // Ler o arquivo `.docx`
        const reader = new FileReader();
        reader.readAsArrayBuffer(modeloFile);

        reader.onload = async function(event) {
            const content = event.target.result;

            // Converter `.docx` para texto mantendo a formatação
            const extractedText = await extractTextFromDocx(content);

            // Substituir as variáveis no texto extraído
            const finalText = extractedText
                .replace(/{{empresa}}/g, empresa)
                .replace(/{{nota}}/g, nota)
                .replace(/{{ordem}}/g, ordem)
                .replace(/{{data}}/g, data);

            // Criar um novo `.docx` com o texto atualizado
            gerarNovoDocx(finalText);
        };
    });
});

// Função para extrair o texto do `.docx` mantendo formatação básica
async function extractTextFromDocx(arrayBuffer) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsBinaryString(new Blob([arrayBuffer]));

        reader.onload = function(event) {
            const binaryString = event.target.result;
            Mammoth.convertToHtml({ arrayBuffer }).then((result) => {
                resolve(result.value);
            }).catch(reject);
        };
    });
}

// Função para gerar um novo `.docx` mantendo a formatação
function gerarNovoDocx(textoAtualizado) {
    const { Document, Packer, Paragraph, TextRun } = docx;

    const doc = new Document({
        sections: [{
            properties: {},
            children: [new Paragraph({ children: [new TextRun(textoAtualizado)] })]
        }]
    });

    Packer.toBlob(doc).then(blob => {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "documento_gerado.docx";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });
}
