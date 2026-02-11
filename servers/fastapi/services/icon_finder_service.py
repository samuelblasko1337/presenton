import asyncio
import json
import os
import logging
from pathlib import Path

import chromadb
from chromadb.config import Settings
from chromadb.utils.embedding_functions import ONNXMiniLM_L6_V2

log = logging.getLogger(__name__)


class IconFinderService:
    def __init__(self):
        self.collection_name = os.getenv("ICONFINDER_COLLECTION", "icons")

        # Wichtig: Pfade konfigurierbar machen (für Volume-Mount / persistente Caches)
        self.chroma_path = os.getenv("ICONFINDER_CHROMA_PATH", "chroma")
        self.model_path = os.getenv("ICONFINDER_MODEL_PATH", "chroma/models")
        self.icons_json_path = os.getenv("ICONFINDER_ICONS_JSON", "assets/icons.json")

        enable = os.getenv("ENABLE_ICON_FINDER", "").lower() in ("1", "true", "yes")
        disable = os.getenv("DISABLE_ICON_FINDER", "").lower() in ("1", "true", "yes")
        self.disabled = (not enable) or disable

        # Lazy-init state
        self._initialized = False
        self._init_error = None

        # Will be set on init
        self.client = None
        self.embedding_function = None
        self.collection = None

    def _initialize_icons_collection(self):
        # Create dirs
        Path(self.chroma_path).mkdir(parents=True, exist_ok=True)
        Path(self.model_path).mkdir(parents=True, exist_ok=True)

        self.client = chromadb.PersistentClient(
            path=self.chroma_path,
            settings=Settings(anonymized_telemetry=False),
        )

        self.embedding_function = ONNXMiniLM_L6_V2()
        self.embedding_function.DOWNLOAD_PATH = self.model_path

        # Optional: kaputten Partial-Download aufräumen (sonst hängt man gerne fest)
        tar_path = Path(self.model_path) / "onnx.tar.gz"
        if tar_path.exists() and tar_path.stat().st_size < 5_000_000:
            # sehr klein => sehr wahrscheinlich kaputt (heuristisch)
            try:
                tar_path.unlink()
            except Exception:
                pass

        # Der Download ist der Pain-Point -> hier darf er NICHT den ganzen Server killen
        self.embedding_function._download_model_if_not_exists()

        # Collection laden oder erstellen
        try:
            self.collection = self.client.get_collection(
                self.collection_name,
                embedding_function=self.embedding_function,
            )
            return
        except Exception:
            pass

        # Create from icons.json
        with open(self.icons_json_path, "r", encoding="utf-8") as f:
            icons = json.load(f)

        documents = []
        ids = []

        for each in icons.get("icons", []):
            name = each.get("name", "")
            tags = each.get("tags", "")
            if name.split("-")[-1] == "bold":
                documents.append(f"{name} {tags}")
                ids.append(name)

        if not documents:
            # Keine Icons => lege wenigstens ne leere Collection an, damit Query nicht crasht
            self.collection = self.client.create_collection(
                name=self.collection_name,
                embedding_function=self.embedding_function,
                metadata={"hnsw:space": "cosine"},
            )
            return

        self.collection = self.client.create_collection(
            name=self.collection_name,
            embedding_function=self.embedding_function,
            metadata={"hnsw:space": "cosine"},
        )
        self.collection.add(documents=documents, ids=ids)

    def ensure_ready(self) -> bool:
        """Lazy init + fail-open."""
        if self.disabled:
            return False
        if self._initialized:
            return True
        if self._init_error is not None:
            return False

        try:
            log.info("Initializing icons collection (lazy)...")
            self._initialize_icons_collection()
            self._initialized = True
            log.info("Icons collection initialized.")
            return True
        except Exception as e:
            # WICHTIG: Kein Crash mehr beim Server-Start
            log.exception("IconFinder init failed; continuing without icons.")
            self._init_error = e

            # Optional: wenn Download halb fertig ist, beim nächsten Run neu versuchen
            tar_path = Path(self.model_path) / "onnx.tar.gz"
            if tar_path.exists():
                try:
                    tar_path.unlink()
                except Exception:
                    pass

            return False

    async def search_icons(self, query: str, k: int = 1):
        if not self.ensure_ready():
            return []

        try:
            result = await asyncio.to_thread(
                self.collection.query,
                query_texts=[query],
                n_results=k,
            )
            ids = result.get("ids", [[]])[0] or []
            return [f"/static/icons/bold/{each}.svg" for each in ids]
        except Exception:
            # Fail-open auch hier
            log.exception("Icon search failed; returning empty result.")
            return []


ICON_FINDER_SERVICE = IconFinderService()
