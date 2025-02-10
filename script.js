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

        // Ler o arquivo `.docx`
        const reader = new FileReader();
        reader.readAsArrayBuffer(modeloFile);

        reader.onload = function(event) {
            const content = event.target.result;

            // Carregar o template `.docx`
            const zip = new PizZip(content);
            const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

            // Substituir variáveis no documento
            doc.render({
                empresa: empresa,
                nota: nota,
                ordem: ordem,
                data: data
            });

            // Gerar novo `.docx`
            const blob = doc.getZip().generate({ type: "blob" });

            // Criar link para download
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = "documento_gerado.docx";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        };
    });
});


