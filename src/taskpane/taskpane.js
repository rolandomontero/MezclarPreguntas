/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// Variables globales para el formulario
let evaluationData = {
  name: '',
  objectives: '',
  course: '',
  date: '',
  unit: '',
  topic: '',
  version: 1
};

/**
 * Inicialización del panel cuando Office está listo
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Ocultar mensaje de sideload y mostrar formulario
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";
    
    // Establecer fecha por defecto (hoy)
    const today = new Date().toISOString().split('T')[0];
    document.getElementById("eval-date").value = today;
    
    // Configurar event listeners
    setupEventListeners();
    
    console.log("Panel de evaluación CSJ inicializado correctamente");
  }
});

/**
 * Configura los event listeners para los botones y campos del formulario
 */
function setupEventListeners() {
  // Botón Crear Versión 1
  const btnCreate = document.getElementById("btn-create-eval");
  if (btnCreate) {
    btnCreate.addEventListener("click", handleCreateEvaluation);
  } else {
    console.error("No se encontró el botón btn-create-eval");
  }
  
  // Botón Cancelar
  const btnCancel = document.getElementById("btn-cancel");
  if (btnCancel) {
    btnCancel.addEventListener("click", handleCancel);
  } else {
    console.error("No se encontró el botón btn-cancel");
  }
  
  // Listeners para validación en tiempo real
  const formFields = [
    'eval-name',
    'eval-objectives',
    'eval-course',
    'eval-date',
    'eval-unit',
    'eval-topic'
  ];
  
  formFields.forEach(fieldId => {
    const field = document.getElementById(fieldId);
    if (field) {
      field.addEventListener('input', () => {
        // Limpiar mensajes de error cuando el usuario escribe
        hideMessage();
      });
      field.addEventListener('blur', () => {
        // Validar campo al perder foco
        validateField(fieldId);
      });
    } else {
      console.error(`No se encontró el campo ${fieldId}`);
    }
  });
}

/**
 * Valida un campo individual del formulario
 * @param {string} fieldId - ID del campo a validar
 * @returns {boolean} - true si es válido, false si hay error
 */
function validateField(fieldId) {
  const field = document.getElementById(fieldId);
  if (!field) return true;
  
  const value = field.value.trim();
  
  // Todos los campos son requeridos
  if (value === '') {
    field.style.borderColor = '#a80000';
    return false;
  } else {
    field.style.borderColor = '#8a8886';
    return true;
  }
}

/**
 * Valida todos los campos del formulario
 * @returns {boolean} - true si todos los campos son válidos
 */
function validateForm() {
  const fields = [
    'eval-name',
    'eval-objectives',
    'eval-course',
    'eval-date',
    'eval-unit',
    'eval-topic'
  ];
  
  let isValid = true;
  
  fields.forEach(fieldId => {
    if (!validateField(fieldId)) {
      isValid = false;
    }
  });
  
  return isValid;
}

/**
 * Recopila los datos del formulario
 * @returns {Object} - Objeto con los datos de la evaluación
 */
function getFormData() {
  return {
    name: document.getElementById("eval-name").value.trim(),
    objectives: document.getElementById("eval-objectives").value.trim(),
    course: document.getElementById("eval-course").value.trim(),
    date: document.getElementById("eval-date").value,
    unit: document.getElementById("eval-unit").value.trim(),
    topic: document.getElementById("eval-topic").value.trim(),
    version: 1
  };
}

/**
 * Maneja el evento de crear evaluación
 */
async function handleCreateEvaluation() {
  console.log("Iniciando creación de evaluación...");
  
  // Limpiar mensajes previos
  hideMessage();
  
  // Validar formulario
  if (!validateForm()) {
    console.log("Validación fallida");
    showMessage(
      "Por favor complete todos los campos requeridos.",
      "error"
    );
    return;
  }
  
  console.log("Validación exitosa, obteniendo datos...");
  
  // Obtener datos del formulario
  evaluationData = getFormData();
  console.log("Datos obtenidos:", evaluationData);
  
  // Mostrar estado de carga
  const btnCreate = document.getElementById("btn-create-eval");
  const originalText = btnCreate.innerHTML;
  btnCreate.disabled = true;
  btnCreate.innerHTML = '<span class="ms-Button-label"><i class="ms-Icon ms-Icon--Refresh"></i> Procesando...</span>';
  
  try {
    console.log("Llamando a createEvaluationFromTemplate...");
    // Llamar a la función que crea el documento desde la plantilla
    await createEvaluationFromTemplate(evaluationData);
    
    console.log("Evaluación creada exitosamente");
    
    // Mostrar mensaje de éxito
    showMessage(
      `Evaluación "${evaluationData.name}" creada exitosamente.`,
      "success"
    );
    
  } catch (error) {
    console.error("Error al crear evaluación:", error);
    console.error("Detalle del error:", error.message);
    console.error("Stack trace:", error.stack);
    showMessage(
      "Error al crear la evaluación: " + error.message,
      "error"
    );
  } finally {
    // Restaurar botón
    btnCreate.disabled = false;
    btnCreate.innerHTML = originalText;
  }
}

/**
 * Crea una nueva evaluación basada en la plantilla PLANTILLA EVALUACION CSJ.dotx
 * @param {Object} data - Datos de la evaluación
 */
async function createEvaluationFromTemplate(data) {
  console.log("Iniciando creación de documento Word...");
  
  return Word.run(async (context) => {
    try {
      console.log("Dentro de Word.run, insertando contenido...");
      
      // Insertar encabezado con información de la evaluación
      const headerParagraph = context.document.body.insertParagraph(
        `EVALUACIÓN: ${data.name}`,
        Word.InsertLocation.start
      );
      headerParagraph.font.bold = true;
      headerParagraph.font.size = 16;
      headerParagraph.alignment = Word.Alignment.center;
      
      console.log("Encabezado insertado");
      
      // Insertar línea en blanco después del encabezado
      const blankLine1 = headerParagraph.insertParagraph(
        "",
        Word.InsertLocation.end
      );
      
      // Insertar información detallada
      const infoText = 
        `Curso: ${data.course}\n` +
        `Unidad: ${data.unit}\n` +
        `Tema: ${data.topic}\n` +
        `Fecha: ${formatDate(data.date)}\n` +
        `Versión: ${data.version}`;
      
      const infoParagraph = headerParagraph.insertParagraph(
        infoText,
        Word.InsertLocation.end
      );
      infoParagraph.font.size = 12;
      
      console.log("Información detallada insertada");
      
      // Insertar línea en blanco después de la información
      const blankLine2 = infoParagraph.insertParagraph(
        "",
        Word.InsertLocation.end
      );
      
      // Insertar objetivos
      const objectivesTitle = infoParagraph.insertParagraph(
        "Objetivos de Aprendizaje:",
        Word.InsertLocation.end
      );
      objectivesTitle.font.bold = true;
      objectivesTitle.font.size = 14;
      
      const objectivesParagraph = objectivesTitle.insertParagraph(
        data.objectives,
        Word.InsertLocation.end
      );
      objectivesParagraph.font.size = 12;
      
      console.log("Objetivos insertados");
      
      // Insertar separador
      const blankLine3 = objectivesParagraph.insertParagraph(
        "",
        Word.InsertLocation.end
      );
      
      const separatorParagraph = objectivesParagraph.insertParagraph(
        "────────────────────────────────────────",
        Word.InsertLocation.end
      );
      separatorParagraph.font.size = 10;
      
      // Insertar instrucciones para preguntas
      const blankLine4 = separatorParagraph.insertParagraph(
        "",
        Word.InsertLocation.end
      );
      
      const questionsInstruction = separatorParagraph.insertParagraph(
        "[Las preguntas se insertarán aquí en la próxima versión]",
        Word.InsertLocation.end
      );
      questionsInstruction.font.italic = true;
      questionsInstruction.font.color = "#666666";
      
      console.log("Contenido completo insertado, sincronizando...");
      
      await context.sync();
      
      console.log("Documento Word creado exitosamente");
      
    } catch (error) {
      console.error("Error dentro de Word.run:", error);
      console.error("Nombre del error:", error.name);
      console.error("Mensaje del error:", error.message);
      console.error("Stack trace:", error.stack);
      throw error;
    }
  });
}

/**
 * Formatea una fecha de YYYY-MM-DD a DD/MM/YYYY
 * @param {string} dateString - Fecha en formato YYYY-MM-DD
 * @returns {string} - Fecha formateada
 */
function formatDate(dateString) {
  if (!dateString) return '';
  const [year, month, day] = dateString.split('-');
  return `${day}/${month}/${year}`;
}

/**
 * Maneja el evento de cancelar
 */
function handleCancel() {
  // Resetear formulario
  document.getElementById("eval-name").value = '';
  document.getElementById("eval-objectives").value = '';
  document.getElementById("eval-course").value = '';
  document.getElementById("eval-unit").value = '';
  document.getElementById("eval-topic").value = '';
  
  // Restablecer fecha a hoy
  const today = new Date().toISOString().split('T')[0];
  document.getElementById("eval-date").value = today;
  
  // Limpiar estilos de validación
  const formFields = document.querySelectorAll('.ms-TextField-field');
  formFields.forEach(field => {
    field.style.borderColor = '#8a8886';
  });
  
  // Ocultar mensajes
  hideMessage();
  
  console.log("Formulario reseteado");
}

/**
 * Muestra un mensaje en el área de mensajes
 * @param {string} text - Texto del mensaje
 * @param {string} type - Tipo de mensaje: 'success', 'error', o 'info'
 */
function showMessage(text, type = 'info') {
  const messageArea = document.getElementById("message-area");
  const messageContent = document.getElementById("message-content");
  
  messageContent.textContent = text;
  messageContent.className = 'ms-Callout-main';
  messageContent.classList.add(`message-${type}`);
  
  messageArea.style.display = 'block';
}

/**
 * Oculta el área de mensajes
 */
function hideMessage() {
  const messageArea = document.getElementById("message-area");
  messageArea.style.display = 'none';
}
