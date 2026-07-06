import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "C:/Users/Eduardo/Desktop/Auditoria/Sistema_de_ventas_Auditoria";
const outDir = path.join(root, "documentacion", "entregables");
await fs.mkdir(outDir, { recursive: true });

const risks = [
  ["Sistema web", "Acceso no autorizado", "Credenciales débiles y contraseñas predecibles", 5, 5, "Crítico", "Política de contraseñas, bloqueo por intentos fallidos y revisión de usuarios."],
  ["Sistema web", "Ingreso sin autenticación válida", "SQL Injection en formulario de login", 5, 3, "Alto", "Consultas parametrizadas, validación de entradas y pruebas de seguridad."],
  ["Sistema web", "Robo de sesión o ejecución de scripts", "XSS reflejado en búsqueda", 3, 3, "Medio", "Escapar salidas HTML, sanitizar entradas y aplicar Content Security Policy."],
  ["Sistema web", "Ejecución persistente de JavaScript", "XSS almacenado en descripción de productos", 5, 3, "Alto", "No usar contenido del usuario como HTML seguro; sanitizar y validar campos."],
  ["Base de datos", "Exposición de credenciales", "Contraseñas almacenadas en texto plano", 5, 3, "Alto", "Hash seguro con bcrypt o argon2 y sal única por usuario."],
  ["Servidor Fedora", "Enumeración de servicios", "Puerto 5000 expuesto durante el laboratorio", 3, 5, "Alto", "Firewall, apertura temporal del puerto y cierre al finalizar la presentación."],
  ["Servidor Fedora", "Resolución de nombres innecesaria", "Puerto 5355/LLMNR abierto", 3, 3, "Medio", "Deshabilitar LLMNR si no es requerido para el laboratorio."],
  ["Aplicación Flask", "Exposición de información técnica", "Modo debug activo", 5, 3, "Alto", "Desactivar debug fuera del laboratorio y usar servidor WSGI para producción."],
  ["Aplicación Flask", "Uso de servidor no productivo", "Werkzeug Development Server expuesto", 3, 5, "Alto", "Usar Gunicorn/uWSGI detrás de Nginx o Apache en entornos reales."],
  ["Servidor Fedora", "Divulgación de información temporal", "ICMP Timestamp Request habilitado", 2, 3, "Bajo", "Filtrar o deshabilitar respuestas ICMP timestamp si no son necesarias."],
  ["Sistema web", "Debilidad en manejo de sesión", "Cookies expiradas o sin endurecimiento suficiente", 3, 3, "Medio", "Configurar atributos Secure, HttpOnly, SameSite y expiración adecuada."],
  ["Sistema web", "Acciones no autorizadas", "Formularios sin token CSRF", 3, 3, "Medio", "Implementar CSRF tokens en formularios y validar origen de solicitudes."],
  ["Servidor Fedora", "Uso indebido de privilegios", "Usuarios o servicios con permisos excesivos", 5, 3, "Alto", "Principio de mínimo privilegio y cuentas separadas para administración."],
  ["Sistema web y servidor", "Pérdida de disponibilidad", "Ausencia de procedimiento de respaldo", 5, 2, "Medio", "Copias periódicas de la base SQLite y del código fuente versionado."],
];

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Matriz de Riesgos");
sheet.showGridLines = false;

sheet.getRange("A1:H1").merge();
sheet.getRange("A1").values = [["Matriz de Análisis de Riesgos - Sistema de Ventas"]];
sheet.getRange("A1").format = {
  fill: "#0F4C5C",
  font: { bold: true, color: "#FFFFFF", size: 14 },
  horizontalAlignment: "center",
};

sheet.getRange("A3:H3").values = [[
  "Activo",
  "Amenaza",
  "Vulnerabilidad",
  "Impacto",
  "Probabilidad",
  "Puntaje",
  "Nivel",
  "Control recomendado",
]];
sheet.getRange("A3:H3").format = {
  fill: "#E8EEF5",
  font: { bold: true, color: "#0B2545" },
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#C8D1DC" },
};

const rows = risks.map((row) => [...row.slice(0, 5), null, row[5], row[6]]);
sheet.getRangeByIndexes(3, 0, rows.length, 8).values = rows;
sheet.getRange(`F4:F${3 + rows.length}`).formulas = risks.map((_, i) => [`=D${4 + i}*E${4 + i}`]);

sheet.getRange(`A4:H${3 + rows.length}`).format = {
  wrapText: true,
  verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: "#D9E2EC" },
};
sheet.getRange(`D4:F${3 + rows.length}`).format = { horizontalAlignment: "center" };
sheet.getRange(`D4:F${3 + rows.length}`).format.numberFormat = "0";

sheet.getRange("A:H").format.font = { name: "Calibri", size: 11 };
sheet.getRange("A:A").format.columnWidth = 20;
sheet.getRange("B:B").format.columnWidth = 28;
sheet.getRange("C:C").format.columnWidth = 34;
sheet.getRange("D:F").format.columnWidth = 12;
sheet.getRange("G:G").format.columnWidth = 14;
sheet.getRange("H:H").format.columnWidth = 46;
sheet.freezePanes.freezeRows(3);

const summary = workbook.worksheets.add("Resumen");
summary.showGridLines = false;
summary.getRange("A1:D1").merge();
summary.getRange("A1").values = [["Resumen de Riesgos"]];
summary.getRange("A1").format = {
  fill: "#0F4C5C",
  font: { bold: true, color: "#FFFFFF", size: 14 },
  horizontalAlignment: "center",
};
summary.getRange("A3:D3").values = [["Nivel", "Cantidad", "Interpretación", "Acción sugerida"]];
summary.getRange("A3:D3").format = {
  fill: "#E8EEF5",
  font: { bold: true, color: "#0B2545" },
  borders: { preset: "all", style: "thin", color: "#C8D1DC" },
};
summary.getRange("A4:D7").values = [
  ["Crítico", null, "Requiere atención inmediata", "Corregir antes de exponer el sistema."],
  ["Alto", null, "Riesgo importante", "Planificar mitigación prioritaria."],
  ["Medio", null, "Riesgo controlable", "Aplicar controles y monitorear."],
  ["Bajo", null, "Riesgo aceptable", "Mantener seguimiento."],
];
summary.getRange("B4:B7").formulas = [
  [`=COUNTIF('Matriz de Riesgos'!G4:G${3 + rows.length},A4)`],
  [`=COUNTIF('Matriz de Riesgos'!G4:G${3 + rows.length},A5)`],
  [`=COUNTIF('Matriz de Riesgos'!G4:G${3 + rows.length},A6)`],
  [`=COUNTIF('Matriz de Riesgos'!G4:G${3 + rows.length},A7)`],
];
summary.getRange("A4:D7").format = {
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#D9E2EC" },
};
summary.getRange("A:D").format.font = { name: "Calibri", size: 11 };
summary.getRange("A:A").format.columnWidth = 16;
summary.getRange("B:B").format.columnWidth = 12;
summary.getRange("C:D").format.columnWidth = 32;

const criteria = workbook.worksheets.add("Criterios");
criteria.showGridLines = false;
criteria.getRange("A1:C1").merge();
criteria.getRange("A1").values = [["Criterios usados para impacto y probabilidad"]];
criteria.getRange("A1").format = {
  fill: "#0F4C5C",
  font: { bold: true, color: "#FFFFFF", size: 14 },
  horizontalAlignment: "center",
};
criteria.getRange("A3:C3").values = [["Valor", "Impacto", "Probabilidad"]];
criteria.getRange("A3:C3").format = {
  fill: "#E8EEF5",
  font: { bold: true, color: "#0B2545" },
  borders: { preset: "all", style: "thin", color: "#C8D1DC" },
};
criteria.getRange("A4:C8").values = [
  [1, "Impacto muy bajo: efecto menor o fácilmente recuperable.", "Muy improbable: requiere condiciones poco realistas."],
  [2, "Impacto bajo: afecta parcialmente una función.", "Poco probable: podría ocurrir, pero no es habitual."],
  [3, "Impacto medio: afecta datos o disponibilidad de forma limitada.", "Probable: puede ocurrir durante el uso normal o con pruebas básicas."],
  [4, "Impacto alto: compromete funciones importantes del sistema.", "Muy probable: puede repetirse con facilidad."],
  [5, "Impacto crítico: compromete acceso, datos sensibles o continuidad.", "Casi seguro: se explota con pasos simples o credenciales conocidas."],
];
criteria.getRange("A4:C8").format = {
  wrapText: true,
  verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: "#D9E2EC" },
};
criteria.getRange("A:C").format.font = { name: "Calibri", size: 11 };
criteria.getRange("A:A").format.columnWidth = 12;
criteria.getRange("B:C").format.columnWidth = 48;

const sources = workbook.worksheets.add("Fuentes");
sources.showGridLines = false;
sources.getRange("A1:B1").values = [["Fuente", "URL"]];
sources.getRange("A1:B1").format = {
  fill: "#E8EEF5",
  font: { bold: true, color: "#0B2545" },
  borders: { preset: "all", style: "thin", color: "#C8D1DC" },
};
sources.getRange("A2:B9").values = [
  ["OWASP Top 10", "https://owasp.org/Top10/2021/"],
  ["OWASP Password Storage Cheat Sheet", "https://cheatsheetseries.owasp.org/cheatsheets/Password_Storage_Cheat_Sheet.html"],
  ["NIST SP 800-63B", "https://csrc.nist.gov/pubs/sp/800/63/b/upd2/final"],
  ["Nmap Reference Guide", "https://nmap.org/book/man.html"],
  ["Tenable Nessus Documentation", "https://docs.tenable.com/Nessus.htm"],
  ["Fedora Server Documentation", "https://docs.fedoraproject.org/en-US/fedora-server/"],
  ["Flask Deployment Documentation", "https://flask.palletsprojects.com/en/stable/deploying/"],
  ["SQLite Appropriate Uses", "https://sqlite.org/whentouse.html"],
];
sources.getRange("A2:B9").format = {
  wrapText: true,
  borders: { preset: "all", style: "thin", color: "#D9E2EC" },
};
sources.getRange("A:A").format.columnWidth = 34;
sources.getRange("B:B").format.columnWidth = 78;

const preview = await workbook.render({ sheetName: "Matriz de Riesgos", autoCrop: "all", scale: 1, format: "png" });
await fs.writeFile(path.join(outDir, "Analisis_Riesgos_preview.png"), new Uint8Array(await preview.arrayBuffer()));

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(path.join(outDir, "Analisis_Riesgos.xlsx"));
