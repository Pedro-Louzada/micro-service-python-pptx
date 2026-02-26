from fastapi import APIRouter, UploadFile, File, Form
from services.squad_generator import SquadProposalGenerator
from services.sustentation_generator import SustentationProposalGenerator
from services.agent_and_sustentation_generator import AgentAndSustentationProposalGenerator
from services.graph_service import GraphService
import json
import logging

logger = logging.getLogger(__name__)

proposal_router = APIRouter(prefix="/proposal", tags=["proposal"])

@proposal_router.post("/generate")
async def generate_proposal(
    payload: str = Form(...),
    logo: UploadFile = File(...)
    ):

    try:
        data = json.loads(payload)
        tipoProposta = data.get("tipoProposta")
        
        logger.info(f"Iniciando geração de proposta: tipo={tipoProposta}")

        match tipoProposta:
            case "SQUAD":
                generator = SquadProposalGenerator("templates/squad.pptx")
                file_path = await generator.generate(data, logo)
                return {"file": file_path}
            
            case "SUSTENTACAO":
                generator = SustentationProposalGenerator("templates/sustentacao.pptx")
                file_path = await generator.generate(data, logo)
                return {"file": file_path}
            
            case "AI AGENT/SUSTENTACAO":
                generator = AgentAndSustentationProposalGenerator("templates/ai-agent-e-sustentacao.pptx")
                file_path = await generator.generate(data, logo)
                return {"file": file_path}

            case "TESTE":
                generator = GraphService()
                file_path = generator.generate(data)
                return {"file": file_path}
            
            case _:
                logger.error(f"Tipo de proposta inválido: {tipoProposta}")
                return {"error": f"Tipo de proposta inválido: {tipoProposta}"}
    
    except json.JSONDecodeError as e:
        logger.error(f"Erro ao fazer parse do payload JSON: {e}")
        return {"error": "Payload JSON inválido"}
    except KeyError as e:
        logger.error(f"Campo obrigatório faltando: {e}")
        return {"error": f"Campo obrigatório faltando: {e}"}
    except Exception as e:
        logger.error(f"Erro ao gerar proposta: {e}", exc_info=True)
        return {"error": f"Erro ao gerar proposta: {str(e)}"}