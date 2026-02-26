from pptx import Presentation
from io import BytesIO
from enum import Enum
from typing import TypedDict
from copy import deepcopy
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
import logging
import io
import re

logger = logging.getLogger(__name__)

class PLANS(Enum):
    STARTER_PLAN = "starter"
    SILVER_PLAN = "silver"
    GOLD_PLAN = "gold"
    DIAMOND_PLAN = "diamond"

class Timeline(TypedDict):
    flowDrawing: str
    drawingHomologation: str
    development: str
    qaHomologation: str
    clientHomologation: str

class AdequatePlanPayload(TypedDict):
    mainGoal: str
    briefingDetails: list[str]
    timeLine: Timeline
    adequatePlan: PLANS

class Client(TypedDict):
    name: str
    briefing: AdequatePlanPayload

class ServiceData(TypedDict):
    tipoProposta: str
    cliente: Client


class AgentAndSustentationProposalGenerator:

    def __init__(self, template_path: str):
        self.prs = Presentation(template_path)

    async def generate(self, data: ServiceData, logo):
        await self._update_logo(logo)
        self._handle_project_scope(data["cliente"]["briefing"])
        self._handle_project_timeline(data["cliente"]["briefing"])
        self._handle_sustentation_plan(data["cliente"]["briefing"])

        output_path = "output/proposta_agent_sustentacao.pptx"
        self.prs.save(output_path)
        return output_path

    async def _update_logo(self, logo_file):
        logger.info(f"Iniciando a atualização da logo...")


        image_bytes = await logo_file.read()
        image_stream = BytesIO(image_bytes)

        slide = self.prs.slides[0]

        for shape in slide.shapes:
            if shape.name == "CLIENT_LOGO":
                left = shape.left
                top = shape.top
                width = shape.width
                height = shape.height

                slide.shapes._spTree.remove(shape._element)

                slide.shapes.add_picture(
                    image_stream,
                    left,
                    top,
                    width=width,
                    height=height
                )
    
    def _normalized_briefing_details(self, detail: str) -> str:
        return f"- {detail.strip()}"
    
    def _chunk_briefing(self, details: list[str], limit: int = 500):
        chunks = []
        current_chunk = ""
        
        for detail in details:
            formatted = f"{detail.strip()}\n"
            
            if len(current_chunk) + len(formatted) > limit:
                chunks.append(current_chunk.strip())
                current_chunk = formatted
            else:
                current_chunk += formatted

        if current_chunk:
            chunks.append(current_chunk.strip())

        return chunks

    def _duplicate_slide(self, slide):
        slide_layout = next(
            layout for layout in self.prs.slide_layouts
            if layout.name.lower() == "blank"
        )

        new_slide = self.prs.slides.add_slide(slide_layout)

        # remover placeholders herdados
        for shape in list(new_slide.shapes):
            if shape.is_placeholder:
                sp = shape._element
                sp.getparent().remove(sp)

        IMAGE_SHAPES = {"PINK_IMAGE", "DIGITALBOT_LOGO"}
        RECTANGLE_SHAPES = {"SCOPE", "SCOPE_MAIN_GOAL", "SCOPE_DETAILS"}

        for shape in slide.shapes:
            # 🔹 Imagens
            if shape.name in IMAGE_SHAPES:
                image_stream = io.BytesIO(shape.image.blob)

                new_picture = new_slide.shapes.add_picture(
                    image_stream,
                    shape.left,
                    shape.top,
                    shape.width,
                    shape.height
                )

                new_picture.name = shape.name

            # 🔹 Retângulos vazios
            elif shape.name in RECTANGLE_SHAPES:
                # Clonar o shape original inteiro com todas as propriedades
                new_element = deepcopy(shape._element)
                new_slide.shapes._spTree.insert_element_before(new_element, 'p:extLst')
                
                if shape.name != "SCOPE":
                    # Encontrar o shape adicionado e limpar seu texto
                    for new_shape in new_slide.shapes:
                        if new_shape.name == shape.name:
                            new_shape.text_frame.clear()
                            break

        # 🔹 Reordenar slide
        slide_id_list = self.prs.slides._sldIdLst

        for idx, sldId in enumerate(slide_id_list):
            if sldId.id == slide.slide_id:
                original_index = idx
                break

        new_sldId = slide_id_list[-1]
        slide_id_list.remove(new_sldId)
        slide_id_list.insert(original_index + 1, new_sldId)

        return new_slide
    
    def _handle_project_scope(self, briefing):
        logger.info(f"Iniciando a atualização dos slides de escopo...")


        main_goal = briefing.get("mainGoal")
        briefing_details = briefing.get("briefingDetails")

        if not briefing_details:
            return

        chunks = self._chunk_briefing(briefing_details, 500)

        scope_slide = None

        for slide in self.prs.slides:
            for shape in slide.shapes:
                if shape.name == "SCOPE":
                    scope_slide = slide
                    break
            if scope_slide:
                break
        
        if not scope_slide:
            return

        slides_to_fill = [scope_slide]

        for _ in range(len(chunks) - 1):
            duplicated = self._duplicate_slide(scope_slide)
            slides_to_fill.append(duplicated)

        for slide, chunk in zip(slides_to_fill, chunks):

            for shape in slide.shapes:
                if shape.name == "SCOPE_MAIN_GOAL":
                    text_frame = shape.text_frame
                    text_frame.clear()

                    p = text_frame.paragraphs[0]
                    p.alignment = PP_ALIGN.LEFT
                    run = p.add_run()
                    run.text = f"🎯 Objetivo Geral: {main_goal}"
                    run.font.name = "Lexend"
                    run.font.size = Pt(22)
                    run.font.color.rgb = RGBColor(0, 0, 0)

                if shape.name == "SCOPE_DETAILS":
                    text_frame = shape.text_frame
                    text_frame.clear()

                    lines = [l.strip().lstrip("-").strip() 
                     for l in chunk.split("\n") if l.strip()]

                    if not lines:
                        continue
                    
                    first_p = text_frame.paragraphs[0]
                    first_p.text = f"• {lines[0]}"
                    first_p.level = 0
                    first_p.alignment = PP_ALIGN.LEFT
                    first_p.runs[0].font.name = "Lexend"
                    first_p.runs[0].font.size = Pt(18)
                    first_p.runs[0].font.color.rgb = RGBColor(0, 0, 0)

                    space_p = text_frame.add_paragraph()
                    space_p.text = ""

                    for line in lines[1:]:
                        p = text_frame.add_paragraph()
                        p.text = f"• {line}"
                        p.level = 0
                        p.alignment = PP_ALIGN.LEFT
                        p.runs[0].font.name = "Lexend"
                        p.runs[0].font.size = Pt(18)
                        p.runs[0].font.color.rgb = RGBColor(0, 0, 0)
                        
                        # Parágrafo vazio entre linhas
                        space_p = text_frame.add_paragraph()
                        space_p.text = ""

    def week_extract(self, texto):
        """Extrai o número de semanas de strings como '2 semanas' ou '1 semana'"""
        if not texto: return 0
        match = re.search(r'(\d+)', str(texto))
        return int(match.group(1)) if match else 0

    def format_transparence_table(self, table, etapas):

        def set_cell_border(cell, color_str="DCDCDC"):
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            
            for side in ['lnL', 'lnR', 'lnT', 'lnB']:
                ln = OxmlElement(f'a:{side}')
                ln.set('w', '6350')  # 0.5pt
                ln.set('cap', 'flat')
                ln.set('cmpd', 'sng')
                ln.set('algn', 'ctr')

                srgbClr = OxmlElement('a:srgbClr')
                srgbClr.set('val', color_str)
                ln.append(srgbClr)

                prstDash = OxmlElement('a:prstDash')
                prstDash.set('val', 'solid')
                ln.append(prstDash)

                tcPr.append(ln)

        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):

                # 🔥 Fundo branco explícito (não background)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

                set_cell_border(cell, "DCDCDC")

                # Primeira coluna com cor da etapa
                if col_idx == 0 and row_idx > 0:
                    cor_etapa = etapas[row_idx-1][2]
                    p = cell.text_frame.paragraphs[0]
                    p.font.name = 'Roboto'
                    p.font.size = Pt(10)
                    p.font.bold = True
                    p.font.color.rgb = cor_etapa
    
    def add_personalized_bar(self, slide, table_shape, row_idx, col_idx, cor_rgb):
        """
        Calcula a posição usando o shape da tabela (GraphicFrame) para evitar erros de atributo.
        """
        table = table_shape.table
        
        # 1. Posição inicial baseada no Shape que contém a tabela
        left_pos = table_shape.left 
        top_pos = table_shape.top
            
        # 2. Somar larguras das colunas anteriores
        for c in range(col_idx):
            left_pos += table.columns[c].width
            
        # 3. Somar alturas das linhas anteriores
        for r in range(row_idx):
            top_pos += table.rows[r].height
            
        # 4. Dimensões da célula alvo
        cell_width = table.columns[col_idx].width
        cell_height = table.rows[row_idx].height
        
        # 5. Padding (Ajuste conforme o gosto visual)
        padding_h = Pt(8) # Padding horizontal
        padding_v = Pt(6) # Padding vertical
        
        bar_left = left_pos + padding_h
        bar_top = top_pos + padding_v
        bar_width = cell_width - (padding_h * 2)
        bar_height = cell_height - (padding_v * 2)
        
        # 6. Adicionar o retângulo arredondado
        rect = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, bar_left, bar_top, bar_width, bar_height
        )
        
        # Estilização final
        rect.fill.solid()
        rect.fill.fore_color.rgb = cor_rgb
        rect.line.fill.background() # Remove contorno
        rect.adjustments[0] = 0.2    # Suaviza o arredondamento

    def _handle_project_timeline(self, briefing: AdequatePlanPayload):
        logger.info(f"Iniciando a construção do gráfico de timeline...")

        timeline_data = briefing["timeLine"]

        graph_shape_name = 'GRAPH_SHAPE'

        etapas = [
            ("flowDrawing", "DESENHO DE FLUXO", RGBColor(31, 73, 125)),
            ("drawingHomologation", "HOMOLOGAÇÃO DO DESENHO", RGBColor(255, 51, 153)),
            ("development", "DESENVOLVIMENTO DO FLUXO", RGBColor(128, 110, 180)),
            ("qaHomologation", "TESTES INTERNOS (QAs)", RGBColor(0, 0, 255)),
            ("clientHomologation", "HOMOLOGAÇÃO CLIENTE", RGBColor(0, 176, 240))
        ]

        for slide in self.prs.slides:
            for shape in slide.shapes:
                if (shape.name == graph_shape_name):
                    total_semanas = sum(self.week_extract(timeline_data.get(k, 0)) for k, _, _ in etapas)
                    num_colunas = max(total_semanas, 6) + 1 

                    rows, cols = len(etapas) + 1, num_colunas
                    
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height

                    graphic_frame = slide.shapes.add_table(rows, cols, left, top, width, height)
                    table = graphic_frame.table
                    table.style = "TableGrid"

                    sp = shape._element
                    sp.getparent().remove(sp)

                    table.first_row = False
                    table.first_col = False
                    table.last_row = False
                    table.last_col = False
                    table.horz_banding = False
                    table.vert_banding = False

                    tbl = table._tbl
                    tblPr = tbl.tblPr
                    tblStyle = tblPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}tblStyle')

                    if tblStyle is not None:
                        tblPr.remove(tblStyle)
                                                            
                    # Limpar fundo e bordas antes de escrever
                    self.format_transparence_table(table, etapas)
                    
                    # 3. Re-inserir Cabeçalhos (SEM 1, SEM 2...)
                    for i in range(1, cols):
                        cell = table.cell(0, i)
                        cell.text = f"SEM {i}"
                        p = cell.text_frame.paragraphs[0]
                        p.alignment = PP_ALIGN.CENTER
                        # Forçar visibilidade e fonte
                        run = p.runs[0]
                        run.font.name = 'Roboto'
                        run.font.size = Pt(10)
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(120, 120, 120) # Cinza discreto para os cabeçalhos

                    # 4. Loop de Preenchimento das Etapas
                    coluna_atual = 1
                    for row_idx, (chave_json, nome_etapa, cor) in enumerate(etapas, start=1):
                        cell_nome = table.cell(row_idx, 0)
                        cell_nome.text = nome_etapa
                        
                        # Formatação Roboto para as nomenclaturas laterais
                        p = cell_nome.text_frame.paragraphs[0]
                        p.alignment = PP_ALIGN.LEFT
                        run = p.runs[0]
                        run.font.name = 'Roboto'
                        run.font.size = Pt(10)
                        run.font.bold = True
                        run.font.color.rgb = cor # Cor correspondente à etapa

                        duracao = self.week_extract(timeline_data.get(chave_json, 0))
                        
                        for _ in range(duracao):
                            if coluna_atual < cols:
                                self.add_personalized_bar(slide, graphic_frame, row_idx, coluna_atual, cor)
                                coluna_atual += 1
                

    def _handle_sustentation_plan(self, briefing: AdequatePlanPayload):
        logger.info(f"Iniciando a escolha de slide de plano de sustentação...")

        valid_plans = {plan.name for plan in PLANS}

        adequate_plan = briefing.get("adequatePlan")

        if not adequate_plan:
            return

        adequate_plan = adequate_plan.upper() + "_PLAN"

        slides_to_remove = []

        for slide in self.prs.slides:
            for shape in slide.shapes:
                if shape.name in valid_plans and shape.name != adequate_plan:
                    slides_to_remove.append(slide)
                    break

        for slide in slides_to_remove:
            self._remove_slide(slide)

    def _remove_slide(self, slide):
        slide_id = slide.slide_id
        slides = self.prs.slides._sldIdLst
        for sld in slides:
            if int(sld.get("id")) == slide_id:
                slides.remove(sld)
                break

