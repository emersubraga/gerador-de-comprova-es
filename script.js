async function carregarMammoth() {
    if (typeof Mammoth === "undefined") {
        return new Promise((resolve, reject) => {
            const script = document.createElement("script");
            script.src = "https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.4.2/mammoth.browser.min.js";
            script.onload = () => {
                console.log("Mammoth.js carregado!");
                resolve();
            };
            script.onerror = () => reject(new Error("Erro ao carregar Mammoth.js"));
            document.head.appendChild(script);
        });
    }
}

await carregarMammoth(); // Garante que Mammoth está carregado antes de ser usado

// Agora podemos usar Mammoth com segurança
const result = await Mammoth.extractRawText({ arrayBuffer });



document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("form").addEventListener("submit", async (event) => {
        event.preventDefault();

        const empresa = document.getElementById("empresa").value.toUpperCase();
        const nota = document.getElementById("nota").value.toUpperCase();
        const ordem = document.getElementById("ordem").value.toUpperCase();
        const data = document.getElementById("data").value.toUpperCase();
        const modeloFile = document.getElementById("modelo").files[0];
        const zipFile = document.getElementById("zip").files[0];

        if (!modeloFile || !zipFile) {
            alert("Por favor, selecione um modelo e um arquivo ZIP.");
            return;
        }

        const modeloArrayBuffer = await modeloFile.arrayBuffer();
        const zip = await JSZip.loadAsync(zipFile);
        const imagens = [];

        for (const fileName of Object.keys(zip.files)) {
            if (/\.(png|jpg|jpeg)$/i.test(fileName)) {
                const imgData = await zip.files[fileName].async("uint8array");
                imagens.push(imgData);
            }
        }

        gerarDocumento(modeloArrayBuffer, empresa, nota, ordem, data, imagens);
    });
});

async function gerarDocumento(modeloArrayBuffer, empresa, nota, ordem, data, imagens) {
    // Converter ArrayBuffer para Blob para usar com Mammoth.js
    const modeloBlob = new Blob([modeloArrayBuffer], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

    // Usando Mammoth.js para extrair o texto do modelo .docx
    const reader = new FileReader();
    reader.readAsArrayBuffer(modeloBlob);
    reader.onload = async (event) => {
        const arrayBuffer = event.target.result;
        const result = await Mammoth.extractRawText({ arrayBuffer });

        // Substituir placeholders no texto extraído
        let textoFormatado = result.value
            .replace("{{empresa}}", empresa)
            .replace("{{nota}}", nota)
            .replace("{{ordem}}", ordem)
            .replace("{{data}}", data);

        // Criar um novo documento com docx.js
        const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun } = docx;
        const doc = new Document({
            sections: [{
                properties: {},
                children: [
                    new Paragraph({ children: [new TextRun(textoFormatado)] }),
                    ...criarTabelaImagens(imagens)
                ]
            }]
        });

        // Gerar e baixar o novo documento
        const blob = await Packer.toBlob(doc);
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "documento_gerado.docx";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };
}

// Função para criar uma tabela de 2x2 com 4 imagens por página
function criarTabelaImagens(imagens) {
    const { Table, TableRow, TableCell, ImageRun } = docx;
    const tabelas = [];
    for (let i = 0; i < imagens.length; i += 4) {
        const imageGroup = imagens.slice(i, i + 4);
        const rows = [];
        for (let j = 0; j < 2; j++) {
            const cells = [];
            for (let k = 0; k < 2; k++) {
                const index = j * 2 + k;
                if (index < imageGroup.length) {
                    cells.push(new TableCell({
                        children: [new Paragraph({ children: [new ImageRun({
                            data: imageGroup[index],
                            transformation: { width: 200, height: 150 }
                        })] })]
                    }));
                } else {
                    cells.push(new TableCell({ children: [] }));
                }
            }
            rows.push(new TableRow({ children: cells }));
        }
        tabelas.push(new Table({ rows }));
    }
    return tabelas;
}


