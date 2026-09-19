/* global Office, Word */

Office.onReady(() => {
    // Nada especial aquí por ahora
});

function onMixQuestions(event) {
    // Ejecuta la misma lógica que el botón del panel
    Word.run(async (context) => {
        // Llamamos a la función del taskpane si está cargada
        try {
            if (window && typeof window.shuffleQuestions === "function") {
                await window.shuffleQuestions();
            }
        } catch (e) {
            console.log("No se pudo llamar a shuffleQuestions desde commands.js:", e);
        }
        event.completed();
    });
}

// Exportar la función para que el manifest la encuentre
if (typeof module !== "undefined" && module.exports) {
    module.exports = {
        onMixQuestions
    };
}
