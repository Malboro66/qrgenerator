from models.lote import ItemLote
from services.lote_service import LoteService
from services.lote_store import LoteStore


def test_gerar_ident_casos():
    assert LoteService.gerar_ident(1, 1, 2026, "011", "122") == "L010126011122"
    assert LoteService.gerar_ident(12, 11, 26, "9999", "2026") == "L121126999026"


def test_seq_crud_status(tmp_path):
    store = LoteStore(str(tmp_path / "lotes.db"))
    service = LoteService(store)
    assert service.proximo_seq_lote(1, 26) == 1
    lote = service.criar_lote(1, 2026, "011", "122", [ItemLote("1", "INJETOR", 2, "011")])
    assert store.proximo_seq_lote(1, 26) == 2
    lotes = store.listar_lotes()
    assert len(lotes) == 1
    store.atualizar_status(lote.id, "impresso")
    assert store.obter_lote(lote.id).status == "impresso"
