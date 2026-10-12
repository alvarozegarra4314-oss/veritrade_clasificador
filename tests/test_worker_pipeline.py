import json
import pickle

from src import worker_pipeline


def test_worker_escribe_resultado_y_estado_completado(tmp_path, monkeypatch):
    ruta_raw = tmp_path / "datos.xlsx"
    ruta_maestro = tmp_path / "maestro.xlsx"
    ruta_raw.write_bytes(b"raw")
    ruta_maestro.write_bytes(b"maestro")
    ruta_resultado = tmp_path / "resultado.pkl"
    ruta_estado = tmp_path / "estado.json"
    esperado = {"kpis": {"total": 1}}

    def clasificar(_raw, _maestro, _hoja, progreso_callback):
        progreso_callback("reglas", 1, 1)
        return esperado

    monkeypatch.setattr(worker_pipeline, "_clasificar", clasificar)

    codigo = worker_pipeline.ejecutar_desde_archivos(
        ruta_raw, ruta_maestro, "Veritrade", ruta_resultado, ruta_estado
    )

    assert codigo == 0
    assert json.loads(ruta_estado.read_text(encoding="utf-8"))["state"] == "completed"
    with ruta_resultado.open("rb") as archivo:
        assert pickle.load(archivo) == esperado


def test_worker_publica_error_de_lectura(tmp_path):
    ruta_estado = tmp_path / "estado.json"

    codigo = worker_pipeline.ejecutar_desde_archivos(
        tmp_path / "no-existe.xlsx",
        tmp_path / "maestro.xlsx",
        "Veritrade",
        tmp_path / "resultado.pkl",
        ruta_estado,
    )

    estado = json.loads(ruta_estado.read_text(encoding="utf-8"))
    assert codigo == 1
    assert estado["state"] == "error"
    assert estado["message"].startswith("FileNotFoundError:")