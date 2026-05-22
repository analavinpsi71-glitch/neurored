[NeuroRed_Hoja_de_Ruta_Paciente.md](https://github.com/user-attachments/files/28131901/NeuroRed_Hoja_de_Ruta_Paciente.md)
[NeuroRed_Consentimiento_Informado.md](https://github.com/user-attachments/files/28131900/NeuroRed_Consentimiento_Informado.md)
[product_brief.js](https://github.com/user-attachments/files/28131904/product_brief.js)const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Footer, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

const CYAN = "0097A7";
const CYAN_LIGHT = "E0F7FA";
const ORANGE = "E65100";
const SLATE = "37474F";
const SLATE_LIGHT = "ECEFF1";
const WHITE = "FFFFFF";

const border = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, font: "Arial", size: 32, bold: true, color: SLATE })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: CYAN, space: 4 } },
    children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: CYAN })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: SLATE })]
  });
}

function body(text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 22, color: SLATE })]
  });
}

function bullet(text, bold_prefix = null) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: bold_prefix
      ? [
          new TextRun({ text: bold_prefix + " ", font: "Arial", size: 22, bold: true, color: SLATE }),
          new TextRun({ text, font: "Arial", size: 22, color: SLATE })
        ]
      : [new TextRun({ text, font: "Arial", size: 22, color: SLATE })]
  });
}

function spacer(lines = 1) {
  return new Paragraph({ spacing: { after: 120 * lines }, children: [new TextRun("")] });
}

function highlight(text) {
  return new Paragraph({
    spacing: { before: 160, after: 160 },
    indent: { left: 720 },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: CYAN, space: 8 } },
    children: [new TextRun({ text, font: "Arial", size: 22, italics: true, color: "555555" })]
  });
}

function actorRow(actor, role, color) {
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 2200, type: WidthType.DXA },
        shading: { fill: color, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        children: [new Paragraph({ children: [new TextRun({ text: actor, font: "Arial", size: 22, bold: true, color: WHITE })] })]
      }),
      new TableCell({
        borders,
        width: { size: 7160, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        children: [new Paragraph({ children: [new TextRun({ text: role, font: "Arial", size: 22, color: SLATE })] })]
      })
    ]
  });
}

function phaseRow(phase, tasks, status) {
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 2200, type: WidthType.DXA },
        shading: { fill: CYAN_LIGHT, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        children: [new Paragraph({ children: [new TextRun({ text: phase, font: "Arial", size: 22, bold: true, color: SLATE })] })]
      }),
      new TableCell({
        borders,
        width: { size: 5560, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        children: [new Paragraph({ children: [new TextRun({ text: tasks, font: "Arial", size: 22, color: SLATE })] })]
      }),
      new TableCell({
        borders,
        width: { size: 1600, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: status, font: "Arial", size: 20, bold: true, color: status === "MVP" ? CYAN : status === "Fase 2" ? "FF9800" : "888888" })] })]
      })
    ]
  });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Psiquiatría Integral Ágil — Product Brief v1.0  |  Confidencial  |  ", font: "Arial", size: 18, color: "999999" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "999999" })
          ]
        })]
      })
    },
    children: [

      // COVER
      new Paragraph({
        spacing: { before: 1200, after: 80 },
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "Psiquiatría Integral Ágil", font: "Arial", size: 52, bold: true, color: CYAN })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "Product Brief — Versión 1.0", font: "Arial", size: 26, color: SLATE })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "Plataforma clínica de monitoreo continuo y reclutamiento para investigación en salud mental", font: "Arial", size: 22, italics: true, color: "777777" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 1200 },
        children: [new TextRun({ text: "Confidencial — 2025", font: "Arial", size: 20, color: "AAAAAA" })]
      }),

      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: CYAN, space: 1 } },
        spacing: { after: 400 },
        children: [new TextRun("")]
      }),

      // 1. PROBLEMA
      h2("1. El problema que resolvemos"),
      spacer(),
      body("El reclutamiento para ensayos clínicos en psiquiatría es ineficiente, costoso y sesgado. Los sistemas actuales de screening evalúan la elegibilidad del paciente recién después de captarlo, sin datos clínicos previos ni seguimiento longitudinal."),
      spacer(),
      highlight("Más del 80% de los pacientes que ingresan a plataformas de reclutamiento existentes resultan no elegibles. Ese descarte representa un enorme desperdicio de tiempo y recursos para sponsors, investigadores y pacientes."),
      spacer(),
      body("Al mismo tiempo, los pacientes psiquiátricos carecen de herramientas de seguimiento continuo entre consultas. El médico solo los ve en turnos presenciales, sin visibilidad de la adherencia, el estado emocional o los cambios clínicos entre sesiones."),

      // 2. SOLUCIÓN
      h2("2. La solución"),
      spacer(),
      body("Psiquiatría Integral Ágil es una plataforma PWA (Progressive Web App) de monitoreo clínico continuo con un modelo de negocio basado en la articulación entre práctica clínica e investigación."),
      spacer(),
      body("El sistema invierte el flujo de reclutamiento tradicional: en lugar de captar pacientes y luego evaluarlos, construye perfiles clínicos longitudinales ricos que permiten identificar eligibilidad de forma precisa, consistente y sin sesgo del evaluador."),
      spacer(),
      h3("Lo que hace la plataforma:"),
      bullet("Monitoreo semanal asincrónico mediante checklist adaptativo al perfil de cada paciente."),
      bullet("Escalas diagnósticas digitalizadas (Hamilton, Beck, PANSS, Young, Columbia y otras) con puntuación automática y devolución al paciente."),
      bullet("Sistema de alertas clínicas por adherencia, síntomas extrapiramidales, valores de laboratorio y desregulación emocional."),
      bullet("Evaluación multimodal: análisis de expresión facial y tono de voz durante la administración de escalas, validado por un rater humano calificado."),
      bullet("Receta digital condicional: el paciente accede a la receta solo si completa el checklist semanal, reforzando la adherencia."),
      bullet("Motor de matching con protocolos de investigación: los criterios de inclusión y exclusión se cargan una vez, y la IA compara automáticamente con los perfiles activos."),
      bullet("El médico tratante es siempre el intermediario entre el paciente y cualquier sponsor o protocolo de investigación."),

      // 3. ACTORES
      h2("3. Ecosistema de actores"),
      spacer(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2200, 7160],
        rows: [
          new TableRow({
            children: [
              new TableCell({
                borders,
                width: { size: 2200, type: WidthType.DXA },
                shading: { fill: SLATE, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 140, right: 140 },
                children: [new Paragraph({ children: [new TextRun({ text: "Actor", font: "Arial", size: 22, bold: true, color: WHITE })] })]
              }),
              new TableCell({
                borders,
                width: { size: 7160, type: WidthType.DXA },
                shading: { fill: SLATE, type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 140, right: 140 },
                children: [new Paragraph({ children: [new TextRun({ text: "Rol y valor recibido", font: "Arial", size: 22, bold: true, color: WHITE })] })]
              })
            ]
          }),
          actorRow("Paciente", "Acceso gratuito a seguimiento clínico continuo, escalas digitales, ejercicios de autorregulación y receta digital. No paga.", "1D9E75"),
          actorRow("Médico tratante", "Panel de monitoreo de toda su cartera de pacientes. Alertas automáticas. Herramienta de investigación desde su consultorio.", CYAN),
          actorRow("Rater clínico", "Valida las evaluaciones multimodales antes de que lleguen al médico. Trabaja de forma remota y asincrónica.", "534AB7"),
          actorRow("Sponsor / CRO", "Accede a pacientes pre-evaluados con perfiles longitudinales y consentimiento documentado. Paga por elegibilidad confirmada.", ORANGE),
        ]
      }),

      // 4. MODELO DE NEGOCIO
      h2("4. Modelo de negocio"),
      spacer(),
      body("El modelo es B2B2C: el paciente usa la plataforma de forma gratuita; los sponsors y CROs pagan por acceder a perfiles elegibles pre-calificados para sus protocolos de investigación."),
      spacer(),
      h3("Fuentes de ingreso:"),
      bullet("Por acceso al pipeline de pacientes elegibles:", "Sponsors / CROs:"),
      bullet("Fee por paciente randomizado en un protocolo de investigación.", ""),
      bullet("Suscripción mensual para acceso a matching activo con criterios cargados.", ""),
      bullet("Licencia de la plataforma para centros de investigación o redes de médicos.", "Centros e instituciones:"),
      bullet("Fee por evaluación multimodal con reporte de rater certificado.", "Evaluaciones premium:"),
      spacer(),
      highlight("El diferencial frente a competidores como 'Un ensayo para mí': los pacientes ya tienen un perfil clínico longitudinal antes del contacto con el sponsor. La tasa de elegibilidad proyectada supera el 60%, frente al 20% de las plataformas actuales."),

      // 5. TECNOLOGÍA
      h2("5. Arquitectura tecnológica"),
      spacer(),
      h3("Fase 1 — MVP clínico (prioritario):"),
      bullet("PWA instalable en móvil y accesible desde web."),
      bullet("Base de datos en Supabase con perfiles de pacientes, registros diarios y resultados de escalas."),
      bullet("Checklist semanal adaptativo y sistema de alertas automáticas."),
      bullet("Escalas digitalizadas con puntuación automática."),
      bullet("Panel del médico con vista de cartera completa."),
      bullet("Receta digital condicional a adherencia."),
      spacer(),
      h3("Fase 2 — Evaluación multimodal:"),
      bullet("Captura de cámara y micrófono durante administración de escalas."),
      bullet("Análisis de expresión facial (API de visión por computadora)."),
      bullet("Análisis de prosodia y tono de voz."),
      bullet("Flujo de consentimiento informado digital por evaluación."),
      bullet("Panel del rater para validación asincrónica."),
      spacer(),
      h3("Fase 3 — Marketplace de protocolos:"),
      bullet("Motor de criterios de inclusión/exclusión cargados por el sponsor."),
      bullet("IA que compara perfiles activos con criterios y genera reportes de elegibilidad."),
      bullet("Panel del sponsor con métricas de pipeline (sin datos personales del paciente)."),
      bullet("Exportación de datos a Excel para análisis de investigación."),

      // 6. ROADMAP
      h2("6. Roadmap"),
      spacer(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2200, 5560, 1600],
        rows: [
          new TableRow({
            children: [
              new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, shading: { fill: SLATE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 140, right: 140 },
                children: [new Paragraph({ children: [new TextRun({ text: "Fase", font: "Arial", size: 22, bold: true, color: WHITE })] })] }),
              new TableCell({ borders, width: { size: 5560, type: WidthType.DXA }, shading: { fill: SLATE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 140, right: 140 },
                children: [new Paragraph({ children: [new TextRun({ text: "Entregables clave", font: "Arial", size: 22, bold: true, color: WHITE })] })] }),
              new TableCell({ borders, width: { size: 1600, type: WidthType.DXA }, shading: { fill: SLATE, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 140, right: 140 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Estado", font: "Arial", size: 22, bold: true, color: WHITE })] })] }),
            ]
          }),
          phaseRow("Fase 1", "App clínica funcional con pacientes reales. Perfiles, escalas, alertas, receta digital, panel médico.", "MVP"),
          phaseRow("Fase 2", "Evaluación multimodal, consentimiento digital, panel rater, análisis facial y de voz.", "Fase 2"),
          phaseRow("Fase 3", "Marketplace de protocolos, matching con IA, panel sponsor, exportación de datos.", "Fase 3"),
        ]
      }),

      // 7. MARCO LEGAL
      h2("7. Marco legal y ético"),
      spacer(),
      bullet("Consentimiento informado digital obligatorio antes de cada evaluación multimodal."),
      bullet("El médico tratante es el intermediario entre el paciente y cualquier sponsor. El sponsor nunca accede a datos personales identificables."),
      bullet("Cumplimiento con Ley 25.326 de Protección de Datos Personales (Argentina) y guías ANMAT para ensayos clínicos."),
      bullet("Revisión legal por especialista en salud digital antes del lanzamiento de la Fase 2."),
      bullet("Todos los datos biométricos se almacenan encriptados y son accesibles solo por el rater autorizado y el médico tratante."),

      // 8. PRÓXIMOS PASOS
      h2("8. Próximos pasos"),
      spacer(),
      bullet("Construir y lanzar el MVP clínico (Fase 1) con la cartera de pacientes actual."),
      bullet("Generar datos reales de adherencia, escalas y alertas durante 60-90 días."),
      bullet("Usar esos datos para el pitch al primer sponsor o CRO interesado en psiquiatría."),
      bullet("Identificar un investigador principal aliado para el primer protocolo piloto."),
      bullet("Incorporar raters clínicos calificados para la Fase 2."),
      spacer(),
      highlight("El objetivo de la Fase 1 no es escalar — es demostrar que el modelo funciona con datos reales. Eso es lo que abre las puertas."),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/mnt/user-data/outputs/Psiquiatria_Integral_Agil_ProductBrief.docx', buffer);
  console.log('Done');
});
