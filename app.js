document.getElementById("generate").addEventListener("click", async () => {
    const excelFile = document.getElementById("excelFile").files[0];

    if (!excelFile) {
        alert("Por favor, selecciona un archivo Excel.");
        return;
    }

    const zip = new JSZip(); // Crear un archivo ZIP para guardar los PDFs

    // Leer el archivo Excel con SheetJS
    const fileReader = new FileReader();
    fileReader.onload = async (event) => {
        const data = event.target.result;
        const workbook = XLSX.read(data, { type: "binary" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet);

        const { jsPDF } = window.jspdf;

        
                // Función para justificar texto
                const addJustifiedText = (doc, text, x, y, maxWidth, lineHeight) => {
                    const words = text.split(" ");
                    let line = "";
                    const spaceWidth = doc.getTextWidth(" ");
        
                    words.forEach((word) => {
                        const testLine = line + word + " ";
                        const testWidth = doc.getTextWidth(testLine);
        
                        if (testWidth > maxWidth) {
                            const gaps = line.split(" ").length - 1;
                            const extraSpace = (maxWidth - doc.getTextWidth(line.trim())) / gaps;
        
                            let currentX = x;
                            line.split(" ").forEach((w, i) => {
                                doc.text(w, currentX, y);
                                currentX += doc.getTextWidth(w) + spaceWidth + (i < gaps ? extraSpace : 0);
                            });
        
                            y += lineHeight;
                            line = word + " ";
                        } else {
                            line = testLine;
                        }
                    });
        
                    if (line.trim()) {
                        doc.text(line.trim(), x, y);
                    }
                };
        

        // Generar PDFs para cada fila
        for (const row of rows) {
            const pdf = new jsPDF();

            // Extraer los datos de la fila
            const folio = row["Folio"];
            const fechaCartaRaw = new Date(); // Obtiene la fecha actual
            const contrato = row["Contrato"];
            const fechaOtorgamiento = row["Fecha de activacion"];
            const nombreAlumno = row["Nombre Alumno"];
            const nombreArchivo = row["Nombre del archivo"];
            const fechaCarta = fechaCartaRaw.toLocaleDateString("es-ES", {
                day: "2-digit",
                month: "2-digit",
                year: "numeric",
            });


            // Agregar logo
            const logo = await fetch("logo.png")
                .then((res) => res.blob())
                .then((blob) => URL.createObjectURL(blob));
            pdf.addImage(logo, "PNG", 10, 8, 50, 25); // Ajustar tamaño y posición del logo

            // Agregar folio y fecha
            pdf.setFontSize(12);
            pdf.text(`Folio: ${folio}`, 190, 10, { align: "right" });
            pdf.text(`Ciudad de México a ${fechaCarta}`, 190, 15, { align: "right" });

            // Título
            pdf.setFontSize(14);
            pdf.text("CARTA FINIQUITO", 105, 67, { align: "center" });

            pdf.setFontSize(12);
            pdf.text(`Suscriptor: ${nombreAlumno}`, 10, 80, { align: "left"});

            // Contenido
            // Contenido
            pdf.setFontSize(12);
            const texto = `Por medio del presente, CORPORATIVO LAUDEX S.A.P.I DE C.V SOFOM E.N.R., hace de su conocimiento la liquidación total del contrato ${contrato} que fue otorgado el ${fechaOtorgamiento} al ciudadano ${nombreAlumno} no ejerciendo alguna responsabilidad, acción y/o derecho entre ambas partes sea de carácter civil, mercantil u otro medio legal, para todos los efectos legales a que haya lugar.`;
            addJustifiedText(pdf, texto, 10, 95, 180, 8);

            const texto2 = `Así mismo se enviará constancia de su comportamiento a las sociedades de información crediticia que corresponda. Se extiende la presente a solicitud del acreditado señalado, con fines informativos y sin responsabilidad alguna para CORPORATIVO LAUDEX S.A.P.I DE C.V SOFOM E.N.R.`;
            addJustifiedText(pdf, texto2, 10, 140, 180, 8);


            // Firma
            const firma = await fetch("firma.png")
                .then((res) => res.blob())
                .then((blob) => URL.createObjectURL(blob));
            pdf.addImage(firma, "PNG", 80, 180, 50, 20);
            pdf.text("Ana Lucia Carbajal", 105, 210, { align: "center" });
            pdf.text("Gerente de Atención a Clientes", 105, 220, { align: "center" });

            // Convertir el PDF a blob y agregarlo al ZIP
            const pdfBlob = pdf.output("blob");
            zip.file(`${nombreArchivo}.pdf`, pdfBlob);
        }

        // Descargar el archivo ZIP con todos los PDFs
        zip.generateAsync({ type: "blob" }).then((content) => {
            const link = document.createElement("a");
            link.href = URL.createObjectURL(content);
            link.download = "CartasFiniquito.zip";
            link.click();
        });

        alert("Cartas generadas correctamente.");
    };

    fileReader.readAsBinaryString(excelFile);
});