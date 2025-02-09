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

        // Ler o modelo .docx
        const modeloArrayBuffer = await modeloFile.arrayBuffer();
        const zip = await JSZip.loadAsync(zipFile);
        const imagens = [];

        // Extrair imagens do ZIP
        for (const fileName of Object.keys(zip.files)) {
            if (/\.(png|jpg|jpeg)$/i.test(fileName)) {
                const imgData = await zip.files[fileName].async("base64");
                imagens.push(`data:image/png;base64,${imgData}`);
            }
        }

        gerarDocumento(modeloArrayBuffer, empresa, nota, ordem, data, imagens);
    });
});

async function gerarDocumento(modeloArrayBuffer, empresa, nota, ordem, data, imagens) {
    const { Document, Packer, Paragraph, TextRun, ImageRun } = docx;

    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun(`Empresa: ${empresa}`).bold(),
                        new TextRun(`\nNota Fiscal: ${nota}`),
                        new TextRun(`\nOrdem de C/S: ${ordem}`),
                        new TextRun(`\nFaturamento Mês/Ano: ${data}`)
                    ]
                }),
                ...imagens.map(imgSrc => new Paragraph({
                    children: [new ImageRun({
                        data: imgSrc,
                        transformation: { width: 400, height: 200 }
                    })]
                }))
            ]
        }]
    });

    // Gerar e baixar o arquivo
    const blob = await Packer.toBlob(doc);
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "documento_gerado.docx";
    link.click();
}

