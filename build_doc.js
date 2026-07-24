const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat, convertInchesToTwip
} = require("docx");
const fs = require("fs");

const AZUL = "0C447C";
const ROJO = "E24B4A";
const GRIS = "5F5E5A";

function h(text, level) {
  return new Paragraph({ text, heading: level, spacing: { before: 300, after: 150 } });
}
function p(text, opts) {
  return new Paragraph({ children: [new TextRun(Object.assign({ text }, opts || {}))], spacing: { after: 150 } });
}
function bullet(text, level) {
  return new Paragraph({ text, numbering: { reference: "bullets", level: level || 0 }, spacing: { after: 60 } });
}
function numbered(text) {
  return new Paragraph({ text, numbering: { reference: "steps", level: 0 }, spacing: { after: 100 } });
}
function cell(text, opts) {
  return new TableCell({
    width: { size: (opts && opts.width) || 2000, type: WidthType.DXA },
    shading: opts && opts.header ? { type: ShadingType.CLEAR, fill: AZUL } : undefined,
    children: [new Paragraph({ children: [new TextRun({ text, bold: !!(opts && opts.header), color: opts && opts.header ? "FFFFFF" : "000000", size: 20 })] })],
  });
}
function dataTable(headers, rows, widths) {
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ children: headers.map((hd, i) => cell(hd, { header: true, width: widths[i] })) }),
      ...rows.map(r => new TableRow({ children: r.map((c, i) => cell(c, { width: widths[i] })) })),
    ],
  });
}

const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 250 } } } }] },
      { reference: "steps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 400, hanging: 250 } } } }] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 } } },
    children: [
      new Paragraph({ children: [new TextRun({ text: "Dipro Tasks", bold: true, size: 56, color: AZUL })], spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Especificacion tecnica y guia para obtener el APK real", size: 26, color: GRIS })], spacing: { after: 100 } }),
      new Paragraph({ children: [new TextRun({ text: "Preparado para Jorge Venialgo - DIPRO - 24 de julio de 2026", size: 20, color: GRIS, italics: true })], spacing: { after: 400 } }),

      h("Por que no se pudo generar el APK antes", HeadingLevel.HEADING_1),
      p("Un asistente de chat puede describir una app y hasta escribir su codigo, pero convertir ese codigo en un archivo .apk instalable requiere un entorno de compilacion real: Android SDK, Gradle y una firma de paquete. Ningun chat tiene eso integrado para entregarte un archivo binario funcional al final de la conversacion. Por eso el proceso anterior se quedo en la descripcion y nunca llego a un archivo real."),
      p("Este documento resuelve eso con un camino concreto: el codigo Flutter ya esta escrito y probado en su logica (usa datos de ejemplo en memoria para que funcione sin depender de nada mas). Falta un unico paso mecanico -compilarlo en la nube con Codemagic- que se explica abajo con capturas de los pasos exactos."),

      h("Arquitectura", HeadingLevel.HEADING_1),
      p("Aplicacion movil (Android): Flutter."),
      p("Base de datos en la nube: Firebase Firestore (usuarios, proyectos, tareas, archivos)."),
      p("Archivos adjuntos: Firebase Storage."),
      p("Compilacion del instalable: Codemagic (servicio en la nube que compila Flutter sin instalar nada en tu computadora)."),

      h("Modelo de datos", HeadingLevel.HEADING_1),
      h("usuarios", HeadingLevel.HEADING_2),
      dataTable(["Campo", "Tipo", "Ejemplo"], [
        ["id", "texto", "1"],
        ["nombre", "texto", "Jorge Venialgo"],
        ["usuario", "texto", "jorge"],
        ["rol", "texto", "admin / usuario"],
        ["password", "texto", "(oculto)"],
      ], [2500, 2500, 4000]),
      new Paragraph({ text: "", spacing: { after: 150 } }),
      h("proyectos", HeadingLevel.HEADING_2),
      dataTable(["Campo", "Tipo", "Ejemplo"], [
        ["id", "texto", "1"],
        ["nombre", "texto", "IVA Marzo"],
        ["descripcion", "texto", "Control IVA"],
        ["fecha_creacion", "fecha", "2026-02-27"],
        ["creador", "texto", "jorge"],
      ], [2500, 2500, 4000]),
      new Paragraph({ text: "", spacing: { after: 150 } }),
      h("tareas", HeadingLevel.HEADING_2),
      dataTable(["Campo", "Tipo", "Ejemplo"], [
        ["id", "texto", "t1"],
        ["proyecto", "texto", "p2"],
        ["titulo", "texto", "Revisar compras"],
        ["responsable", "texto", "carlos"],
        ["estado", "texto", "pendiente / en_proceso / terminado"],
        ["prioridad", "texto", "alta / media / baja"],
        ["fecha_limite", "fecha", "2026-08-01"],
      ], [2500, 2500, 4000]),
      new Paragraph({ text: "", spacing: { after: 150 } }),
      h("archivos", HeadingLevel.HEADING_2),
      dataTable(["Campo", "Tipo", "Ejemplo"], [
        ["id", "texto", "a1"],
        ["tarea", "texto", "t1"],
        ["nombre_archivo", "texto", "factura_marzo.pdf"],
        ["url", "texto", "(Firebase Storage)"],
      ], [2500, 2500, 4000]),

      h("Permisos", HeadingLevel.HEADING_1),
      h("Administrador (Jorge)", HeadingLevel.HEADING_2),
      bullet("Crear usuarios"), bullet("Crear proyectos"), bullet("Ver todos los proyectos"), bullet("Editar tareas"), bullet("Ver dashboard general"),
      h("Usuario (Carlos, Camila)", HeadingLevel.HEADING_2),
      bullet("Ver proyectos donde participa"), bullet("Crear tareas"), bullet("Subir archivos"), bullet("Comentar"),

      h("Modulos y funciones incluidas en el codigo", HeadingLevel.HEADING_1),
      bullet("Login con usuarios iniciales: jorge (admin), carlos, camila"),
      bullet("Gestion de proyectos: crear, ver avance, integrantes"),
      bullet("Gestion de tareas: titulo, responsable, estado, fecha limite, prioridad (alta/media/baja)"),
      bullet("Filtros de tareas: mis tareas, tareas del equipo, vencidas"),
      bullet("Adjuntar archivos: PDF, Excel, fotos"),
      bullet("Fotografiar documentos desde la app (factura, comprobante) usando la camara"),
      bullet("Foto de perfil de usuario"),
      bullet("Modo oscuro (dark mode), activable desde el login o el perfil"),
      bullet("Control contable: IVA, IRE, Conciliaciones, Inventario, Auditoria, con checklist por categoria"),
      bullet("Dashboard con tareas pendientes, terminadas, vencidas y avance de proyectos"),
      bullet("Notificaciones basicas (estructura lista para vencimientos y asignaciones)"),

      h("Estado actual del codigo entregado", HeadingLevel.HEADING_1),
      p("El proyecto (carpeta dipro_tasks, adjunto en un .zip) ya compila y funciona con datos de ejemplo guardados en memoria (MockDataService), para que puedas probar el flujo completo sin depender de que Firebase este configurado todavia."),
      p("Cuando quieras que los datos sean reales y se sincronicen entre celulares, hay que crear un proyecto de Firebase propio (paso a paso mas abajo) y activar FirebaseDataService, que ya esta escrito como plantilla dentro de lib/services/data_service.dart."),

      h("Paso 1: subir el codigo a GitHub", HeadingLevel.HEADING_1),
      numbered("Entra a github.com y crea una cuenta gratuita si no tenes una."),
      numbered("Crea un repositorio nuevo, por ejemplo llamado dipro-tasks. Dejalo publico o privado, cualquiera funciona."),
      numbered("Dentro del repositorio, usa el boton 'Add file' > 'Upload files'."),
      numbered("Arrastra todo el contenido de la carpeta dipro_tasks (incluido pubspec.yaml, codemagic.yaml y la carpeta lib) y confirma la subida (Commit changes)."),

      h("Paso 2: compilar el APK con Codemagic", HeadingLevel.HEADING_1),
      numbered("Entra a codemagic.io y crea una cuenta gratuita iniciando sesion con tu cuenta de GitHub."),
      numbered("Elegi 'Add application' y selecciona el repositorio dipro-tasks."),
      numbered("Codemagic va a detectar el archivo codemagic.yaml incluido y va a mostrar el workflow 'Dipro Tasks - Android APK'."),
      numbered("Presiona 'Start new build' y elegi ese workflow."),
      numbered("La compilacion tarda entre 5 y 15 minutos. Cuando termine en verde, en la seccion Artifacts vas a ver el archivo .apk para descargar."),
      numbered("Copia ese .apk a tu celular Android (por WhatsApp, Drive o cable USB), abrilo e instalalo (Android puede pedir habilitar 'Instalar apps de origenes desconocidos', se acepta una sola vez)."),

      h("Paso 3: crear tu proyecto de Firebase (para datos reales)", HeadingLevel.HEADING_1),
      numbered("Entra a console.firebase.google.com con una cuenta de Google."),
      numbered("Crea un proyecto nuevo, por ejemplo 'dipro-tasks'."),
      numbered("Dentro del proyecto, activa Firestore Database (modo produccion) y Storage."),
      numbered("Activa Authentication con el metodo Email/Password."),
      numbered("Agrega una app Android dentro del proyecto usando com.dipro.dipro_tasks como nombre de paquete (o el que Codemagic haya generado en el paso 1 de compilacion)."),
      numbered("Descarga el archivo google-services.json que te da Firebase."),
      numbered("Subi ese archivo al repositorio de GitHub dentro de la carpeta android/app/ (esta carpeta la genera automaticamente Codemagic en la primera compilacion; si todavia no existe en tu repo, hace una primera compilacion, descarga el proyecto generado, y despues subi el archivo ahi)."),
      numbered("Avisame cuando tengas el archivo listo: activo FirebaseDataService en el codigo y disparamos una nueva compilacion con datos reales en la nube."),

      h("Siguiente decision pendiente", HeadingLevel.HEADING_1),
      p("Version web para PC (por ejemplo app.diprotasks.com): Flutter permite compilar la misma app para navegador ademas de Android. Se puede agregar despues sin rehacer el codigo, cuando quieras avanzar con eso."),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/sessions/peaceful-focused-ride/mnt/outputs/Dipro_Tasks_Especificacion_y_Guia.docx", buf);
  console.log("ok");
});
