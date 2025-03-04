/**
 * Script para extraer información de reuniones diarias de Slack a Google Sheets
 * Autor: MahooBox
 * Fecha: Marzo 2025
 * Versión: 1.4 - Se agregó procesamiento avanzado con IA y diccionarios contextuales
 */

// Datos de configuración
const SLACK_OAUTH_TOKEN = 'xoxb-84xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx';
const CHANNEL_ID = 'C0xxxxxxxxxx'; // ID del canal de Slack
const SPREADSHEET_ID = '1jtzwxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'; // ID de Google Sheet
const DAYS_TO_FETCH = 3; // Número de días hacia atrás para consultar
const GEMINI_API_KEY = 'AIzaxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'; 

// Lista de IDs de usuario para ignorar
const IGNORED_USER_IDS = ['USLACKBOT', 'U0000000000']; 


// Diccionario de empresas internas
const INTERNAL_COMPANIES = {
  "empresamatriz": {
    "nombre": "Empresa CORP",
    "especialidad": "Casa matriz de todas las empresas. Desarrollo de tecnologías y soluciones de innovación"
  },
  "empresa1": {
    "nombre": "Company Labs",
    "especialidad": "Laboratorio de innovación y desarrollo de software/hardware para mejorar la Productividad y sistemas agroalimentarios."
  }
};

// Diccionario de proyectos
const PROJECTS = {
  "acuicultura": {
    "nombre": "Acuicultura Fish",
    "convenio": "03334",
    "cliente": "Universidad Mi Universidad",
    "objeto": "Desarrollo de tecnologías para monitoreo y optimización de cultivos acuícolas"
  },
  "manglares": {
    "nombre": "Restauración de Manglares",
    "convenio": "000",
    "cliente": "Corporación Autónoma ",
    "objeto": "Monitoreo y restauración de ecosistemas de manglar con tecnología drone"
  }
};

// Cache para almacenar información de usuarios y evitar llamadas repetidas a la API
let userCache = {};

/**
 * Envía un resumen de bloqueos/impedimentos a un canal de Slack específico
 */
function enviarResumenBloqueosASlack() {
  // Obtener la fecha actual en formato YYYY-MM-DD
  const fechaActual = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  
  // Obtener la hoja con los datos consolidados
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Dailys Consolidados');
  
  if (!sheet) {
    Logger.log("No se encontró la hoja 'Dailys Consolidados'");
    return;
  }
  
  // Obtener todos los datos
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // Encontrar índices de columnas relevantes
  const fechaIndex = headers.indexOf('Fecha');
  const nombreIndex = headers.indexOf('Nombre');
  const bloqueosIndex = headers.indexOf('Bloqueos/Impedimentos');
  
  // Filtrar registros de la fecha actual y con bloqueos
  const registrosHoy = data.slice(1).filter(row => {
    // Convertir la fecha de la celda a formato YYYY-MM-DD para comparar
    const fechaRegistro = Utilities.formatDate(new Date(row[fechaIndex]), "GMT", "yyyy-MM-dd");
    const bloqueos = row[bloqueosIndex];
    
    return fechaRegistro === fechaActual && 
           bloqueos && 
           bloqueos !== "N/A" && 
           bloqueos.toLowerCase() !== "n/a" &&
           bloqueos.toLowerCase() !== "ninguno";
  });
  
  // Crear el mensaje con el formato requerido
  let mensaje = `- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -\nBloqueos / Impedimentos ${fechaActual} \n\n`;
  
  // Agregar bloqueos por usuario
  registrosHoy.forEach(row => {
    const nombre = row[nombreIndex];
    const bloqueos = row[bloqueosIndex];
    
    mensaje += `${nombre}\n${bloqueos}\n\n`;
  });
  
  // Agregar la lista de usuarios que reportaron resultados
  mensaje += "USUARIOS REPORTADOS:\n";
  
  // Obtener lista de usuarios únicos que reportaron hoy
  const usuariosReportaron = data.slice(1)
    .filter(row => {
      const fechaRegistro = Utilities.formatDate(new Date(row[fechaIndex]), "GMT", "yyyy-MM-dd");
      return fechaRegistro === fechaActual;
    });
  
  // Extraer los user_ids y obtener nombres reales desde el cache
  const userIdIndex = headers.indexOf('User ID');
  const usuariosReportados = [...new Set(usuariosReportaron.map(row => {
    const userId = row[userIdIndex];
    return userCache[userId] ? userCache[userId].real_name : row[nombreIndex];
  }))].sort(); // Añadimos .sort() para ordenar alfabéticamente
  
  mensaje += usuariosReportados.join('\n');

  
  // Enviar el mensaje a Slack
  const slackApiUrl = 'https://slack.com/api/chat.postMessage';
  const payload = {
    "channel": "Cxxxxxxx",
    "text": mensaje
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${SLACK_OAUTH_TOKEN}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(slackApiUrl, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.ok) {
      Logger.log("Mensaje de resumen de bloqueos enviado correctamente a Slack");
    } else {
      Logger.log(`Error al enviar mensaje a Slack: ${result.error}`);
    }
  } catch (error) {
    Logger.log(`Excepción al enviar mensaje a Slack: ${error}`);
  }
}

/**
 * Función principal que ejecuta todo el proceso
 */
function procesarDailyUpdates() {
  // Obtener información de usuarios para enriquecer los datos
  fetchAndCacheUsers();
  
  // Obtener mensajes de Slack
  const messages = fetchSlackMessages();
  
  // Procesar mensajes y extraer la información estructurada
  const dailyData = processMessages(messages);
  
  // Escribir datos en Google Sheets (Registro Diario Bruto)
  writeToSheet(dailyData, 'Registro Diario Bruto');
  
  // Ejecutar el procesamiento con Gemini para obtener dailys consolidados
  processWithGemini();
  
  // Crear hojas de seguimiento y dashboard
  createTrackingSheets();

  // Enviar resumen de bloqueos a Slack
  enviarResumenBloqueosASlack();
  
  Logger.log("Se completó la actualización completa del proceso");
}

/**
 * Obtiene y almacena en caché la información de todos los usuarios del workspace
 */
function fetchAndCacheUsers() {
  try {
    const url = 'https://slack.com/api/users.list?limit=100';
    
    const options = {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${SLACK_OAUTH_TOKEN}`,
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      Logger.log(`Error al obtener información de usuarios: ${data.error}`);
      return;
    }
    
    // Guardar información de usuarios en caché
    userCache = {};
    data.members.forEach(member => {
      userCache[member.id] = {
        real_name: member.profile.real_name || member.name || 'Sin nombre',
        email: member.profile.email || 'No disponible',
        is_bot: member.is_bot || false
      };
    });
    
    Logger.log(`Se cargó información de ${Object.keys(userCache).length} usuarios.`);
  } catch (error) {
    Logger.log(`Error en la solicitud a la API de Slack para obtener usuarios: ${error}`);
  }
}

/**
 * Obtiene los mensajes del canal de Slack a través de la API
 * @return {Array} Array de mensajes de Slack
 */
function fetchSlackMessages() {
  // Obtener la fecha actual en formato YYYY-MM-DD
  const fechaActual = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
  
  // Calcular timestamp para el inicio del día actual (en segundos desde la época Unix)
  const startOfDay = new Date(fechaActual);
  const oldestTimestamp = Math.floor(startOfDay.getTime() / 1000);
  
  // Usar solo el parámetro oldest para obtener mensajes desde el inicio del día hasta ahora
  const url = `https://slack.com/api/conversations.history?channel=${CHANNEL_ID}&oldest=${oldestTimestamp}&limit=100`;
  
  Logger.log(`URL de la solicitud: ${url}`);
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${SLACK_OAUTH_TOKEN}`,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      Logger.log(`Error al obtener mensajes de Slack: ${data.error}`);
      return [];
    }
    
    Logger.log(`Se obtuvieron ${data.messages.length} mensajes del día actual (${fechaActual})`);
    
    // Filtrar mensajes de usuarios ignorados
    const filteredMessages = data.messages.filter(msg => 
      !IGNORED_USER_IDS.includes(msg.user)
    );
    
    Logger.log(`Se filtraron ${data.messages.length - filteredMessages.length} mensajes de usuarios ignorados`);
    return filteredMessages;
  } catch (error) {
    Logger.log(`Error en la solicitud a la API de Slack: ${error}`);
    return [];
  }
}

/**
 * Procesa los mensajes de Slack y extrae la información estructurada
 * @param {Array} messages Array de mensajes de Slack
 * @return {Array} Array de objetos con la información diaria estructurada
 */
function processMessages(messages) {
  const dailyData = [];
  
  // Filtrar solo mensajes tipo "message" que no sean subtipos como channel_join
  const actualMessages = messages.filter(msg => 
    msg.type === "message" && 
    !msg.subtype && 
    msg.text && 
    (msg.text.includes("Lo que hice ayer") || 
     msg.text.includes("lo que hice ayer") || 
     msg.text.includes("1. Lo que hice"))
  );
  
  for (const message of actualMessages) {
    try {
      // Extraer fecha del timestamp
      const msgTimestamp = parseFloat(message.ts) * 1000;
      const msgDate = new Date(msgTimestamp).toISOString().split('T')[0];
      
      // Obtener información del usuario
      const userId = message.user || "No disponible";
      const userInfo = userCache[userId] || { 
        real_name: "No disponible", 
        email: "No disponible" 
      };
      
      // Extraer nombre y rol
      let name = userInfo.real_name;
      let role = "No especificado";
      const roleMatch = message.text.match(/Rol:\s*(.*?)(?:\n|$)/i);
      if (roleMatch && roleMatch[1]) {
        role = roleMatch[1].trim();
      }
      
      // Extraer secciones principales
      let yesterday = "";
      let today = "";
      let blockers = "";
      
      // Extraer lo que hizo ayer
      const yesterdayMatch = message.text.match(/(?:1\.\s*)?[Ll]o que hice ayer:?([\s\S]*?)(?:2\.|Lo que haré hoy|$)/i);
      if (yesterdayMatch && yesterdayMatch[1]) {
        yesterday = yesterdayMatch[1].trim().replace(/^\s*-\s*/gm, '').replace(/\n+/g, ' | ');
      }
      
      // Extraer lo que hará hoy
      const todayMatch = message.text.match(/(?:2\.\s*)?[Ll]o que haré hoy:?([\s\S]*?)(?:3\.|Bloqueos|$)/i);
      if (todayMatch && todayMatch[1]) {
        today = todayMatch[1].trim().replace(/^\s*-\s*/gm, '').replace(/\n+/g, ' | ');
      }
      
      // Extraer bloqueos o impedimentos
      const blockersMatch = message.text.match(/(?:3\.\s*)?[Bb]loqueos o impedimentos:?([\s\S]*?)(?:$)/i);
      if (blockersMatch && blockersMatch[1]) {
        blockers = blockersMatch[1].trim().replace(/^\s*-\s*/gm, '').replace(/\n+/g, ' | ');
      }
      
      // Extraer proyectos mencionados
      const projectsRegex = new RegExp(Object.keys(PROJECTS).join('|'), 'gi');
      let projects = "";
      let projectMatches = message.text.match(projectsRegex);
      if (projectMatches) {
        projects = [...new Set(projectMatches.map(p => p.toLowerCase()))].join(', ');
      }
      
      // Estado del bloqueador
      const blockerStatus = blockers && blockers.toLowerCase() !== "n/a" && blockers.toLowerCase() !== "ninguno" ? "Activo" : "Ninguno";
      
      dailyData.push({
        fecha: msgDate,
        user_id: userId,
        nombre: name,
        email: userInfo.email,
        rol: role,
        ayer: yesterday,
        hoy: today,
        bloqueos: blockers,
        estadoBloqueo: blockerStatus,
        proyectos: projects,
        raw_timestamp: message.ts,
        raw_text: message.text // Guardar el texto original para referencia
      });
    } catch (error) {
      Logger.log(`Error procesando mensaje: ${error}`);
      Logger.log(`Mensaje problemático: ${JSON.stringify(message)}`);
    }
  }
  
  // Ordenar por fecha (más reciente primero) y luego por nombre
  return dailyData.sort((a, b) => {
    if (a.fecha !== b.fecha) {
      return b.fecha.localeCompare(a.fecha);
    }
    return a.nombre.localeCompare(b.nombre);
  });
}

/**
 * Escribe los datos procesados en la hoja de cálculo
 * @param {Array} dailyData Array de objetos con la información diaria estructurada
 * @param {string} sheetName Nombre de la hoja donde se escribirán los datos
 */
function writeToSheet(dailyData, sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  // Crear la hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = ['Fecha', 'User ID', 'Nombre', 'Email', 'Rol', 'Ayer (Completado)', 'Hoy (Planificado)', 
                     'Bloqueos/Impedimentos', 'Estado del Bloqueo', 'Proyectos', 'Timestamp'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#E0E0E0');
    sheet.setFrozenRows(1);
  }
  
  // Verificar si ya existen registros con los mismos timestamps para evitar duplicados
  const existingData = sheet.getDataRange().getValues();
  const existingTimestamps = new Set();
  
  // Comenzar desde la fila 1 (que son los encabezados)
  for (let i = 1; i < existingData.length; i++) {
    const timestamp = existingData[i][10]; // Columna de timestamp (índice 10)
    if (timestamp) {
      existingTimestamps.add(timestamp);
    }
  }
  
  // Filtrar solo nuevos mensajes que no existan ya en la hoja
  const newData = dailyData.filter(item => !existingTimestamps.has(item.raw_timestamp));
  
  if (newData.length > 0) {
    // Preparar datos para insertar
    const rowsToAdd = newData.map(item => [
      item.fecha,
      item.user_id,
      item.nombre,
      item.email,
      item.rol,
      item.ayer,
      item.hoy,
      item.bloqueos,
      item.estadoBloqueo,
      item.proyectos,
      item.raw_timestamp
    ]);
    
    // Añadir nuevas filas
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    
    // Dar formato a la hoja
    sheet.autoResizeColumns(1, 11);
    
    Logger.log(`Se agregaron ${rowsToAdd.length} nuevos registros en ${sheetName}.`);
  } else {
    Logger.log(`No se encontraron nuevos registros para agregar en ${sheetName}.`);
  }
}

/**
 * Procesa los mensajes de dailys usando la API de Gemini para extraer información estructurada
 * aun cuando los mensajes no siguen el formato estándar
 */
function processWithGemini() {
  // Asegurar que tengamos la información de usuarios cargada
  fetchAndCacheUsers();
  
  // Obtener los mensajes de Slack
  const messages = fetchSlackMessages();
  
  // Filtrar solo mensajes tipo "message" que son potencialmente dailys
  // con criterios más amplios para capturar formatos no estándar
  const potentialDailyMessages = messages.filter(msg => 
    msg.type === "message" && 
    !msg.subtype && 
    msg.text && 
    msg.text.length > 50 && // Mensajes con cierta longitud
    !IGNORED_USER_IDS.includes(msg.user) // Excluir usuarios ignorados
  );
  
  // Crear array para almacenar los resultados procesados
  const processedDailys = [];
  
  // Procesar cada mensaje con Gemini
  for (const message of potentialDailyMessages) {
    try {
      // Llamar a Gemini para analizar el mensaje
      const processedDaily = analyzeWithGemini(message);
      
      if (processedDaily) {
        processedDailys.push(processedDaily);
      }
    } catch (error) {
      Logger.log(`Error procesando mensaje con Gemini: ${error}`);
      Logger.log(`Mensaje problemático: ${JSON.stringify(message)}`);
    }
  }
  
  // Escribir los resultados en la hoja de Dailys Consolidados
  writeGeminiResultsToSheet(processedDailys);
  
  return processedDailys;
}

/**
 * Analiza un mensaje de Slack usando la API de Gemini
 * @param {Object} message - Mensaje de Slack
 * @return {Object} Información estructurada del daily
 */
function analyzeWithGemini(message) {
  const GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent';
  
  // Extraer información básica del mensaje
  const userId = message.user || "No disponible";
  const timestamp = message.ts;
  const messageDate = new Date(parseFloat(timestamp) * 1000).toISOString().split('T')[0];
  
  // Obtener información del usuario desde la caché
  const userInfo = userCache[userId] || { 
    real_name: "No disponible", 
    email: "No disponible" 
  };
  
  // Preparar contexto con información de proyectos y empresas para la IA
  const contextInfo = {
    proyectos: PROJECTS,
    empresas: INTERNAL_COMPANIES
  };
  
  // Instrucciones para Gemini
  const prompt = `
  Eres un analista de dailys que ayuda a interpretar reportes diarios no estructurados de un equipo de desarrollo.
  
  Analiza el siguiente mensaje y extrae la siguiente información en formato JSON:
  
  1. nombre (string): El nombre de la persona.
  2. rol (string): El rol de la persona, si está disponible.
  3. ayer (array): Lista de tareas completadas ayer, separando claramente cada tarea.
  4. hoy (array): Lista de tareas planificadas para hoy, separando claramente cada tarea.
  5. bloqueos (array): Lista de bloqueos o impedimentos. Usa ["N/A"] si no hay bloqueos.
  6. estadoBloqueo (string): "Activo" si hay bloqueos reales, "Ninguno" si no hay bloqueos.
  7. proyectos (array): Lista de proyectos mencionados. Usa la información del contexto para identificar proyectos.
  
  IMPORTANTE:
  - El mensaje puede no seguir un formato estándar. Analiza todo su contenido para extraer la información.
  - Si hay información sobre lo que la persona hizo recientemente, colócala en "ayer".
  - Si hay información sobre planes o tareas futuras, colócala en "hoy".
  - Si el usuario menciona bloqueos, problemas, conflictos o cualquier item que se pueda entender como necesidades en cualquier parte del texto, captúralos como "bloqueos".
  - Reconoce abreviaturas y jerga técnica del contexto proporcionado.
  - Incluye solo información real, no agregues tareas inventadas.
  - Actúa como un SCRUM Master y Resume cada tarea. 
  
  CONTEXTO (Diccionarios de proyectos y empresas para referencia):
  ${JSON.stringify(contextInfo, null, 2)}
  
  MENSAJE:
  ${message.text}
  
  Responde SOLO con el JSON, sin texto adicional.
  `;
  
  // Datos para la solicitud a la API
  const requestData = {
    "contents": [{
      "parts": [{"text": prompt}]
    }]
  };
  
  // Opciones para la solicitud HTTP
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(requestData),
    'muteHttpExceptions': true
  };
  
  try {
    // Realizar la solicitud a la API de Gemini
    const response = UrlFetchApp.fetch(`${GEMINI_API_URL}?key=${GEMINI_API_KEY}`, options);
    const responseJson = JSON.parse(response.getContentText());
    
    // Verificar que hay una respuesta
    if (responseJson.candidates && responseJson.candidates.length > 0) {
      const content = responseJson.candidates[0].content;
      if (content && content.parts && content.parts.length > 0) {
        // Intentar extraer el JSON de la respuesta
        const jsonText = content.parts[0].text.trim();
        
        // Encontrar el inicio y fin del JSON en caso de que haya texto adicional
        const jsonStart = jsonText.indexOf('{');
        const jsonEnd = jsonText.lastIndexOf('}') + 1;
        
        if (jsonStart >= 0 && jsonEnd > jsonStart) {
          const jsonPortion = jsonText.substring(jsonStart, jsonEnd);
          const parsedJson = JSON.parse(jsonPortion);

          // Procesar arrays para formato de Google Sheets (con bullets y saltos de línea)
          let formattedYesterday = Array.isArray(parsedJson.ayer) 
            ? parsedJson.ayer.map(task => `• ${task}`).join('\n') 
            : (parsedJson.ayer || "");
            
          let formattedToday = Array.isArray(parsedJson.hoy) 
            ? parsedJson.hoy.map(task => `• ${task}`).join('\n') 
            : (parsedJson.hoy || "");
            
          let formattedBlockers = Array.isArray(parsedJson.bloqueos) 
            ? parsedJson.bloqueos.filter(b => b !== "N/A").map(blocker => `• ${blocker}`).join('\n') 
            : (parsedJson.bloqueos || "N/A");
          
          if (formattedBlockers.trim() === "" || formattedBlockers === "• N/A") {
            formattedBlockers = "N/A";
          }
          
          let formattedProjects = Array.isArray(parsedJson.proyectos) 
            ? parsedJson.proyectos.join(', ') 
            : (parsedJson.proyectos || "");
          
          // Usar el nombre real del usuario si Gemini no pudo extraer uno
          const nombreUsuario = userInfo.real_name || parsedJson.nombre || "No disponible";
          
          // Añadir información adicional al objeto
          return {
            fecha: messageDate,
            user_id: userId,
            nombre: userInfo.real_name || parsedJson.nombre || "No disponible",
            email: userInfo.email,
            rol: parsedJson.rol || "No especificado",
            ayer: formattedYesterday,
            hoy: formattedToday,
            bloqueos: formattedBlockers,
            estadoBloqueo: parsedJson.estadoBloqueo || (formattedBlockers === "N/A" ? "Ninguno" : "Activo"),
            proyectos: formattedProjects,
            timestamp: timestamp
          };
        }
      }
    }
    
    Logger.log(`No se pudo extraer JSON de la respuesta de Gemini para el usuario: ${userInfo.real_name}`);
    return null;
  } catch (error) {
    Logger.log(`Error en la solicitud a la API de Gemini: ${error}`);
    return null;
  }
}

/**
 * Escribe los resultados procesados con Gemini en la hoja de Dailys Consolidados
 * @param {Array} processedDailys Array de objetos con la información diaria estructurada
 */
function writeGeminiResultsToSheet(processedDailys) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Dailys Consolidados');
  
  // Crear la hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet('Dailys Consolidados');
    const headers = ['Fecha', 'User ID', 'Nombre', 'Email', 'Rol', 'Ayer (Completado)', 'Hoy (Planificado)', 
                     'Bloqueos/Impedimentos', 'Estado del Bloqueo', 'Proyectos', 'Timestamp'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#B6D7A8');
    sheet.setFrozenRows(1);
    
    // Establecer el formato de fecha para la columna de Fecha
    sheet.getRange("A2:A").setNumberFormat("yyyy-mm-dd");
    
    // Configurar ajuste de texto para facilitar lectura de contenido con múltiples líneas
    sheet.getRange("F:H").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  
  // Verificar si ya existen registros con los mismos timestamps para evitar duplicados
  const existingData = sheet.getDataRange().getValues();
  const existingTimestamps = new Set();
  
  // Comenzar desde la fila 1 (que son los encabezados)
  for (let i = 1; i < existingData.length; i++) {
    const timestamp = existingData[i][10]; // Columna de timestamp (índice 10)
    if (timestamp) {
      existingTimestamps.add(timestamp);
    }
  }

  // Filtrar solo nuevos mensajes que no existan ya en la hoja
  const newData = processedDailys.filter(item => !existingTimestamps.has(item.timestamp));
  
  if (newData.length > 0) {
    // Preparar datos para insertar
    const rowsToAdd = newData.map(item => [
      item.fecha,
      item.user_id,
      item.nombre,
      item.email,
      item.rol,
      item.ayer,
      item.hoy,
      item.bloqueos,
      item.estadoBloqueo,
      item.proyectos,
      item.timestamp
    ]);
    
    // Añadir nuevas filas
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    
    // Dar formato a la hoja
    sheet.autoResizeColumns(1, 5); // Columnas sin texto largo
    sheet.autoResizeColumns(9, 2); // Columnas sin texto largo
    
    // Establecer un ancho fijo para las columnas con texto largo
    sheet.setColumnWidth(6, 400); // Ayer
    sheet.setColumnWidth(7, 400); // Hoy
    sheet.setColumnWidth(8, 400); // Bloqueos
    
    Logger.log(`Se agregaron ${rowsToAdd.length} nuevos registros procesados por Gemini.`);
  } else {
    Logger.log("No se encontraron nuevos registros para agregar.");
  }
}

/**
 * Crea o actualiza las hojas de seguimiento de tareas y bloqueos
 */
function createTrackingSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Crear Hoja de Seguimiento de Tareas
  let taskSheet = ss.getSheetByName('Seguimiento de Tareas');
  if (!taskSheet) {
    taskSheet = ss.insertSheet('Seguimiento de Tareas');
    const taskHeaders = ['Fecha Mencionada', 'User ID', 'Nombre', 'Email', 'Tarea', 'Estado', 'Días en Estado', 'Proyecto', 'Categoría'];
    taskSheet.appendRow(taskHeaders);
    taskSheet.getRange(1, 1, 1, taskHeaders.length).setFontWeight('bold').setBackground('#D9EAD3');
    
    // Fórmula para extraer tareas de la hoja de Dailys Consolidados
    taskSheet.getRange('A2').setFormula('=QUERY(\'Dailys Consolidados\'!A:K, "SELECT A, B, C, D, G, \'Planificada\', DATEDIF(A, TODAY(), \'D\'), J, \'\' WHERE G IS NOT NULL ORDER BY A DESC, C ASC", 1)');
  } else {
    // Actualizar la fórmula en caso de que la estructura haya cambiado
    taskSheet.getRange('A2').setFormula('=QUERY(\'Dailys Consolidados\'!A:K, "SELECT A, B, C, D, G, \'Planificada\', DATEDIF(A, TODAY(), \'D\'), J, \'\' WHERE G IS NOT NULL ORDER BY A DESC, C ASC", 1)');
  }

  // 2. Crear Hoja de Seguimiento de Bloqueos
  let blockersSheet = ss.getSheetByName('Bloqueos');
  if (!blockersSheet) {
    blockersSheet = ss.insertSheet('Bloqueos');
    const blockerHeaders = ['Fecha Reportada', 'User ID', 'Nombre', 'Email', 'Descripción del Bloqueo', 'Estado', 'Días Activo', 'Responsable de Resolución', 'Fecha de Resolución'];
    blockersSheet.appendRow(blockerHeaders);
    blockersSheet.getRange(1, 1, 1, blockerHeaders.length).setFontWeight('bold').setBackground('#F4CCCC');
    
    // Fórmula para extraer bloqueos de la hoja de Dailys Consolidados
    blockersSheet.getRange('A2').setFormula('=QUERY(\'Dailys Consolidados\'!A:K, "SELECT A, B, C, D, H, I, DATEDIF(A, TODAY(), \'D\'), \'\', \'\' WHERE I = \'Activo\' ORDER BY A DESC, C ASC", 1)');
  } else {
    // Actualizar la fórmula en caso de que la estructura haya cambiado
    blockersSheet.getRange('A2').setFormula('=QUERY(\'Dailys Consolidados\'!A:K, "SELECT A, B, C, D, H, I, DATEDIF(A, TODAY(), \'D\'), \'\', \'\' WHERE I = \'Activo\' ORDER BY A DESC, C ASC", 1)');
  }
  
  // 3. Crear Hoja de Participación
  let participationSheet = ss.getSheetByName('Participación');
  if (!participationSheet) {
    participationSheet = ss.insertSheet('Participación');
    const participationHeaders = ['User ID', 'Nombre', 'Email', 'Rol', 'Días Reportados', 'Última Participación', 'Proyectos'];
    participationSheet.appendRow(participationHeaders);
    participationSheet.getRange(1, 1, 1, participationHeaders.length).setFontWeight('bold').setBackground('#D0E0E3');
    
    // Fórmula para extraer participación
    participationSheet.getRange('A2').setFormula('=QUERY(\'Dailys Consolidados\'!A:K, "SELECT B, C, D, E, COUNT(A), MAX(A), J WHERE B IS NOT NULL GROUP BY B, C, D, E, J ORDER BY COUNT(A) DESC", 1)');
  } else {
    // Actualizar la fórmula en caso de que la estructura haya cambiado
    participationSheet.getRange('A2').setFormula('=QUERY(\'Dailys Consolidados\'!A:K, "SELECT B, C, D, E, COUNT(A), MAX(A), J WHERE B IS NOT NULL GROUP BY B, C, D, E, J ORDER BY COUNT(A) DESC", 1)');
  }
  
  // 4. Crear Dashboard
  let dashboardSheet = ss.getSheetByName('Dashboard');
  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet('Dashboard');
    dashboardSheet.getRange('A1').setValue('Dashboard de Seguimiento Daily');
    dashboardSheet.getRange('A1').setFontSize(16).setFontWeight('bold');
    
    // Crear gráficos y tablas resumen
    dashboardSheet.getRange('A3').setValue('Resumen de Actividad por Persona (Últimos 7 días)');
    dashboardSheet.getRange('A3').setFontWeight('bold');
    dashboardSheet.getRange('A4').setFormula('=COUNTIFS(\'Dailys Consolidados\'!A:A,">="&TODAY()-7)');
    dashboardSheet.getRange('A5').setValue('reportes en la última semana');
    
    dashboardSheet.getRange('A7').setValue('Proyectos Activos:');
    dashboardSheet.getRange('A7').setFontWeight('bold');
    dashboardSheet.getRange('A8').setFormula('=QUERY(\'Dailys Consolidados\'!A:K, "SELECT J, COUNT(DISTINCT B) WHERE A >= date \'"&TEXT(TODAY()-7,"yyyy-MM-dd")&"\' GROUP BY J ORDER BY COUNT(DISTINCT B) DESC", 1)');
    
    dashboardSheet.getRange('D3').setValue('Bloqueos Activos:');
    dashboardSheet.getRange('D3').setFontWeight('bold');
    dashboardSheet.getRange('D4').setFormula('=COUNTIFS(\'Dailys Consolidados\'!I:I,"Activo")');
    
    dashboardSheet.getRange('D7').setValue('Miembros del Equipo Activos:');
    dashboardSheet.getRange('D7').setFontWeight('bold');
    dashboardSheet.getRange('D8').setFormula('=QUERY(\'Dailys Consolidados\'!A:K, "SELECT C, D, MAX(A) WHERE A >= date \'"&TEXT(TODAY()-7,"yyyy-MM-dd")&"\' GROUP BY C, D ORDER BY MAX(A) DESC", 1)');
    
    // Formatear dashboard
    dashboardSheet.autoResizeColumns(1, 8);
  } else {
    // Actualizar las fórmulas del dashboard
    dashboardSheet.getRange('A8').setFormula('=QUERY(\'Dailys Consolidados\'!A:K, "SELECT J, COUNT(DISTINCT B) WHERE A >= date \'"&TEXT(TODAY()-7,"yyyy-MM-dd")&"\' GROUP BY J ORDER BY COUNT(DISTINCT B) DESC", 1)');
    dashboardSheet.getRange('D8').setFormula('=QUERY(\'Dailys Consolidados\'!A:K, "SELECT C, D, MAX(A) WHERE A >= date \'"&TEXT(TODAY()-7,"yyyy-MM-dd")&"\' GROUP BY C, D ORDER BY MAX(A) DESC", 1)');
  }
  
  Logger.log("Se crearon/actualizaron las hojas de seguimiento");
}

/**
 * Configura un disparador para ejecutar la función cada día a las 12:00 PM

function setTrigger() {
  // Eliminar triggers existentes para evitar duplicados
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'procesarDailyUpdates') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // Crear nuevo trigger diario para la función principal
  ScriptApp.newTrigger('procesarDailyUpdates')
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();
  
  Logger.log('Trigger configurado para ejecutar el script diariamente a las 12:00 PM');
}
*/

/**
 * Función para crear un menú personalizado en la hoja de cálculo (reducido)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Slack Daily')
    .addItem('Procesar datos de Slack', 'procesarDailyUpdates')
    .addToUi();
}
