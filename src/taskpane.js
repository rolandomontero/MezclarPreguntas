Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        const btn = document.getElementById("shuffle-btn");
        if (btn) {
            btn.onclick = shuffleQuestions;
        }
    }
});

async function shuffleQuestions() {
    const status = document.getElementById("status");
    if (status) status.innerText = "Procesando...";

    try {
        await Word.run(async (context) => {
            const body = context.document.body;
            const paragraphs = body.paragraphs;
            paragraphs.load("text, style");
            await context.sync();

            const styleIds = {
                question: "Preguntadeseleccinmltiple",
                selection: "Seleccin",
                options: "Opciones"
            };

            const normalize = s => (s || "").trim().toLowerCase();

            let questionBlocks = [];
            let currentBlock = null;

            // 1. Agrupar párrafos en bloques de preguntas
            paragraphs.items.forEach((para) => {
                const style = para.style;
                const text = para.text;
                const styleNorm = normalize(style);

                if (styleNorm === normalize(styleIds.question)) {
                    // Si ya había un bloque, lo guardamos
                    if (currentBlock) {
                        questionBlocks.push(currentBlock);
                    }
                    // Iniciamos nuevo bloque
                    currentBlock = {
                        questionText: text,
                        questionStyle: style,
                        options: []
                    };
                } else if (
                    styleNorm === normalize(styleIds.selection) ||
                    styleNorm === normalize(styleIds.options)
                ) {
                    if (currentBlock) {
                        currentBlock.options.push({
                            text: text,
                            style: style
                        });
                    }
                } else {
                    // Otros estilos se ignoran para el mezclado
                }
            });

            // Agregar el último bloque
            if (currentBlock) {
                questionBlocks.push(currentBlock);
            }

            if (questionBlocks.length === 0) {
                throw new Error("No se encontraron párrafos con los estilos requeridos (Preguntadeseleccinmltiple, Seleccin, Opciones).");
            }

            // 2. Mezclar el orden de las preguntas (Fisher-Yates)
            for (let i = questionBlocks.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [questionBlocks[i], questionBlocks[j]] = [questionBlocks[j], questionBlocks[i]];
            }

            // 3. Mezclar las opciones dentro de cada pregunta
            questionBlocks.forEach(block => {
                for (let i = block.options.length - 1; i > 0; i--) {
                    const j = Math.floor(Math.random() * (i + 1));
                    [block.options[i], block.options[j]] = [block.options[j], block.options[i]];
                }
            });

            // 4. Reconstruir el documento
            let finalContent = [];
            questionBlocks.forEach(block => {
                finalContent.push({ text: block.questionText, style: block.questionStyle });
                block.options.forEach(opt => {
                    finalContent.push({ text: opt.text, style: opt.style });
                });
                // Puedes agregar un párrafo en blanco entre preguntas si quieres:
                // finalContent.push({ text: "", style: block.questionStyle });
            });

            // Limpiar el cuerpo del documento
            body.clear();

            // Insertar los párrafos mezclados
            finalContent.forEach(item => {
                const newPara = body.insertParagraph(item.text, Word.InsertLocation.end);
                newPara.style = item.style;
            });

            if (status) status.innerText = "¡Mezclado completado con éxito!";
        });
    } catch (error) {
        console.error(error);
        if (status) status.innerText = "Error: " + error.message;
    }
}
