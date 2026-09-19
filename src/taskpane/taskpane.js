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
  }
  
  // Botón Cancelar
  const btnCancel = document.getElementById("btn-cancel");
  if (btnCancel) {
    btnCancel.addEventListener("click", handleCancel);
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
  // Limpiar mensajes previos
  hideMessage();
  
  // Validar formulario
  if (!validateForm()) {
    showMessage(
      "Por favor complete todos los campos requeridos.",
      "error"
    );
    return;
  }
  
  // Obtener datos del formulario
  evaluationData = getFormData();
  
  // Mostrar estado de carga
  const btnCreate = document.getElementById("btn-create-eval");
  const originalText = btnCreate.innerHTML;
  btnCreate.disabled = true;
  btnCreate.innerHTML = '<span class="ms-Button-label"><i class="ms-Icon ms-Icon--Refresh"></i> Procesando...</span>';
  
  try {
    // Llamar a la función que crea el documento desde la plantilla
    await createEvaluationFromTemplate(evaluationData);
    
    // Mostrar mensaje de éxito
    showMessage(
      `Evaluación "${evaluationData.name}" creada exitosamente.`,
      "success"
    );
    
  } catch (error) {
    console.error("Error al crear evaluación:", error);
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
  return Word.run(async (context) => {
    try {
      // Insertar información de la evaluación en el documento
      // NOTA: Esta es una implementación básica. En versiones futuras se puede:
      // 1. Cargar la plantilla .dotx completa
      // 2. Reemplazar marcadores de posición
      // 3. Mezclar preguntas y alternativas
      
      // Insertar encabezado con información de la evaluación
      const headerParagraph = context.document.body.insertParagraph(
        `EVALUACIÓN: ${data.name}`,
        Word.InsertLocation.start
      );
      headerParagraph.font.bold = true;
      headerParagraph.font.size = 16;
      headerParagraph.alignment = Word.Alignment.center;
      
      // Insertar línea en blanco
      context.document.body.insertParagraph("", Word.InsertLocation.after(headerParagraph));
      
      // Insertar información detallada
      const infoText = 
        `Curso: ${data.course}\n` +
        `Unidad: ${data.unit}\n` +
        `Tema: ${data.topic}\n` +
        `Fecha: ${formatDate(data.date)}\n` +
        `Versión: ${data.version}`;
      
      const infoParagraph = context.document.body.insertParagraph(
        infoText,
        Word.InsertLocation.after(headerParagraph)
      );
      infoParagraph.font.size = 12;
      
      // Insertar objetivos
      context.document.body.insertParagraph("", Word.InsertLocation.after(infoParagraph));
      
      const objectivesTitle = context.document.body.insertParagraph(
        "Objetivos de Aprendizaje:",
        Word.InsertLocation.after(infoParagraph)
      );
      objectivesTitle.font.bold = true;
      objectivesTitle.font.size = 14;
      
      const objectivesParagraph = context.document.body.insertParagraph(
        data.objectives,
        Word.InsertLocation.after(objectivesTitle)
      );
      objectivesParagraph.font.size = 12;
      
      // Insertar separador
      context.document.body.insertParagraph("", Word.InsertLocation.after(objectivesParagraph));
      const separatorParagraph = context.document.body.insertParagraph(
        "────────────────────────────────────────",
        Word.InsertLocation.after(objectivesParagraph)
      );
      separatorParagraph.font.size = 10;
      
      // Insertar instrucciones para preguntas
      context.document.body.insertParagraph("", Word.InsertLocation.after(separatorParagraph));
      const questionsInstruction = context.document.body.insertParagraph(
        "[Las preguntas se insertarán aquí en la próxima versión]",
        Word.InsertLocation.after(separatorParagraph)
      );
      questionsInstruction.font.italic = true;
      questionsInstruction.font.color = "#666666";
      
      await context.sync();
      
      console.log("Evaluación creada exitosamente");
      
    } catch (error) {
      console.error("Error en Word.run:", error);
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
