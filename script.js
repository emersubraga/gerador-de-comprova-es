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
            if (/(.png|.jpg|.jpeg)$/i.test(fileName)) {
                const imgData = await zip.files[fileName].async("uint8array");
                imagens.push(imgData);
            }
        }

        gerarDocumento(modeloArrayBuffer, empresa, nota, ordem, data, imagens);
    });
});

async function gerarDocumento(modeloArrayBuffer, empresa, nota, ordem, data, imagens) {
    const doc = await new docx.DocumentLoader(modeloArrayBuffer).load();

    const paragraphs = doc.getParagraphs();
    for (const paragraph of paragraphs) {
        let text = paragraph.getText();
        text = text.replace("{{empresa}}", empresa)
                   .replace("{{nota}}", nota)
                   .replace("{{ordem}}", ordem)
                   .replace("{{data}}", data);
        paragraph.replaceText(text);
    }

    let imageCounter = 0;
    for (let i = 0; i < imagens.length; i += 4) {
        const imageGroup = imagens.slice(i, i + 4);
        const table = new docx.Table({ rows: 2, columns: 2 });
        imageGroup.forEach((img, index) => {
            table.getCell(Math.floor(index / 2), index % 2).addContent(
                new docx.ImageRun({ data: img, transformation: { width: 200, height: 150 } })
            );
        });
        doc.addSection({ children: [table] });
        imageCounter++;
    }

    const blob = await docx.Packer.toBlob(doc);
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "documento_gerado.docx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

