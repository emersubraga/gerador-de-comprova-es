document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("form").addEventListener("submit", async (event) => {
        event.preventDefault(); // ✅ Impede o reload da página

        console.log("Iniciando geração do documento..."); // ✅ Debug

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

        console.log("Arquivo modelo e ZIP selecionados."); // ✅ Debug

        const modeloArrayBuffer = await modeloFile.arrayBuffer();
        const zip = await JSZip.loadAsync(zipFile);
        const imagens = [];

        for (const fileName of Object.keys(zip.files)) {
            if (/\.(png|jpg|jpeg)$/i.test(fileName)) {
                const imgData = await zip.files[fileName].async("base64");
                imagens.push(`data:image/png;base64,${imgData}`);
            }
        }

        console.log("Imagens extraídas do ZIP:", imagens.length); // ✅ Debug

        gerarDocumento(modeloArrayBuffer, empresa, nota, ordem, data, imagens);
    });
});

async function gerarDocumento(modeloArrayBuffer, empresa, nota, ordem, data, imagens) { 
    console.log("Gerando documento..."); // ✅ Debug

    const { Document, Packer, Paragraph, TextRun, ImageRun } = docx;

    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun({ text: `Empresa: ${empresa}` }),
                        new TextRun({ text: `\nNota Fiscal: ${nota}` }),
                        new TextRun({ text: `\nOrdem de C/S: ${ordem}` }),
                        new TextRun({ text: `\nFaturamento Mês/Ano: ${data}` })
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

    console.log("Documento criado, preparando para download..."); // ✅ Debug

    try {
        const blob = await Packer.toBlob(doc);
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "documento_gerado.docx";
        document.body.appendChild(link); // ✅ Adiciona temporariamente ao DOM
        link.click();
        document.body.removeChild(link); // ✅ Remove o link após o clique

        console.log("Download iniciado."); // ✅ Debug
    } catch (error) {
        console.error("Erro ao gerar documento:", error);
    }
}
