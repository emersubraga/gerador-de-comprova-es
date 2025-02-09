async function gerarDocumento(modeloArrayBuffer, empresa, nota, ordem, data, imagens) {
    const { Document, Packer, Paragraph, TextRun, ImageRun } = docx;

    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun({ text: `Empresa: ${empresa}`, bold: true }),
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


