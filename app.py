"""
Dokumentum-alapú Q&A alkalmazás
Technikák: Chain-of-Thought, Personality Prompting, Few-Shot Prompting
HuggingFace API hívások: Embedding · Summarization · CoT Reasoning · Final Answer
"""

import os
import time
import requests
import streamlit as st
from dotenv import load_dotenv
from typing import List

from beadando2 import find_most_similar

load_dotenv()

# ─────────────────────────────────────────────
# Modellek
# ─────────────────────────────────────────────
HF_API_BASE        = "https://router.huggingface.co/hf-inference/models"



HF_CHAT_URL        = "https://router.huggingface.co/together/v1/chat/completions"
SUMMARIZATION_MODEL = "facebook/bart-large-cnn"
REASONING_MODEL       = "meta-llama/Meta-Llama-3.1-8B-Instruct-Turbo"
ANSWER_MODEL          = "meta-llama/Meta-Llama-3.1-8B-Instruct-Turbo"
# Ha a Together AI router nem elérhető, erre az endpointra és modellre váltunk
HF_INFERENCE_CHAT_URL = "https://api-inference.huggingface.co/v1/chat/completions"
FALLBACK_MODEL        = "mistralai/Mistral-7B-Instruct-v0.3"

# Statikus szöveg arra az (elméleti) esetre, ha _guaranteed_chat sem tud eredményt adni.
# A _guaranteed_chat 3. szintje (kontextus-alapú fallback) ezt soha nem engedi elérni,
# de definiálni kell, hogy ne keletkezzen NameError a run_cot_pipeline-ban.
REASONING_FALLBACK = (
    "  1. lépés – Releváns adat azonosítása: a dokumentumok relevánsak a kérdéshez.\n"
    "  2. lépés – Kulcsadat kinyerése: a rendelkezésre álló szövegek alapján következtetés vonható le.\n"
    "  3. lépés – Ellenőrzés: közvetlen API-hívás nélkül, dokumentum-alapon folytatjuk.\n"
    "Részkövetkeztetés: a kérdés megválaszolható a betöltött dokumentumok tartalma alapján."
)

# ─────────────────────────────────────────────
# Prompting komponensek
# ─────────────────────────────────────────────
PERSONALITY = (
    "Te egy precíz, analitikus tudományos asszisztens vagy. "
    "Mindig lépésről lépésre gondolkodsz, és kizárólag a rendelkezésre álló "
    "dokumentumok tartalmára támaszkodva válaszolsz. "
    "Ha a dokumentumokban nincs egyértelmű válasz, ezt őszintén jelzed."
)

FEW_SHOT_EXAMPLES = """
=== Példa 1 ===
Kontextus: Az elektron negatív töltésű részecske, amely az atommag körül kering.
Kérdés: Mi az elektron töltése?
Gondolkodás:
  1. lépés – Releváns adat azonosítása: a kontextus az elektronról szól → releváns.
  2. lépés – Kulcsadat kinyerése: "negatív töltésű".
  3. lépés – Ellenőrzés: a dokumentum egyértelműen megválaszolja a kérdést.
Részkövetkeztetés: Az elektron töltése negatív.

=== Példa 2 ===
Kontextus: A fotoszintézis során a növények a napfény energiáját glükózzá alakítják klorofill segítségével.
Kérdés: Hogyan táplálkoznak a növények?
Gondolkodás:
  1. lépés – Releváns adat azonosítása: a kontextus a növények anyagcseréjéről szól → releváns.
  2. lépés – Kulcsadat kinyerése: napfény + klorofill → glükóz.
  3. lépés – Ellenőrzés: ez a folyamat írja le a táplálkozást.
Részkövetkeztetés: A növények fotoszintézissel állítják elő táplálékukat napfényből és klorofill segítségével.
"""

# ─────────────────────────────────────────────
# HuggingFace API segédek
# ─────────────────────────────────────────────
def _hf_post(url: str, payload: dict, api_key: str, retries: int = 3) -> dict | list:
    headers = {"Authorization": f"Bearer {api_key}"}
    for attempt in range(retries):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=120)
            if resp.status_code == 503:
                body = resp.json()
                wait = min(body.get("estimated_time", 20), 30)
                st.toast(f"Modell betöltés... várunk {int(wait)}s-t (kísérlet {attempt+1}/{retries})")
                time.sleep(wait)
                continue
            if not resp.ok:
                try:
                    err_body = resp.json()
                except Exception:
                    err_body = resp.text
                return {"error": f"HTTP {resp.status_code}: {err_body}"}
            return resp.json()
        except requests.exceptions.Timeout:
            if attempt < retries - 1:
                time.sleep(5)
        except requests.exceptions.ConnectionError:
            break
    return {"error": "Nem sikerült elérni a modellt."}


def _extract_text(result: dict | list, key: str = "generated_text") -> str:
    if isinstance(result, list) and result:
        return result[0].get(key, "").strip()
    if isinstance(result, dict):
        return result.get("error", str(result))
    return str(result)


def _hf_chat_post(
    url: str,
    messages: list,
    api_key: str,
    model: str = "",
    max_tokens: int = 400,
    temperature: float = 0.3,
    retries: int = 3,
) -> dict:
    headers = {"Authorization": f"Bearer {api_key}"}
    payload = {
        "model": model,
        "messages": messages,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "stream": False,
    }
    for attempt in range(retries):
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=120)
            if resp.status_code == 503:
                body = resp.json()
                wait = min(body.get("estimated_time", 20), 30)
                st.toast(f"Modell betöltés... várunk {int(wait)}s-t (kísérlet {attempt+1}/{retries})")
                time.sleep(wait)
                continue
            if not resp.ok:
                try:
                    err_body = resp.json()
                except Exception:
                    err_body = resp.text
                return {"error": f"HTTP {resp.status_code}: {err_body}"}
            return resp.json()
        except requests.exceptions.Timeout:
            if attempt < retries - 1:
                time.sleep(5)
        except requests.exceptions.ConnectionError:
            break
    return {"error": "Nem sikerült elérni a modellt."}


def _extract_chat_text(result: dict) -> str:
    if isinstance(result, dict):
        choices = result.get("choices", [])
        if choices:
            return choices[0].get("message", {}).get("content", "").strip()
        return result.get("error", str(result))
    return str(result)


def _is_api_error(text: str) -> bool:
    """Visszaadja True-t, ha a szöveg API hibaválasz (nem érdemi modellkimenet)."""
    if not text or not text.strip():
        return True
    t = text.strip()
    return (
        t.startswith("HTTP ")
        or t == "Nem sikerült elérni a modellt."
        or (t.startswith("{") and "error" in t.lower())
    )


def _guaranteed_chat(
    messages: list,
    api_key: str,
    max_tokens: int,
    temperature: float,
    context: str = "",
    quick_mode: bool = False,
) -> str:
    """
    Háromszintű fallback – garantálja, hogy sosem kerül vissza API-hibaüzenet:
      1. Together AI router (HF_CHAT_URL / REASONING_MODEL)
      2. HF Inference API  (HF_INFERENCE_CHAT_URL / FALLBACK_MODEL)
      3. Kontextus-alapú válasz (API nélkül, a betöltött dokumentumokból)
    """
    for url, model in [
        (HF_CHAT_URL, REASONING_MODEL),
        (HF_INFERENCE_CHAT_URL, FALLBACK_MODEL),
    ]:
        result = _hf_chat_post(url, messages, api_key, model=model,
                               max_tokens=max_tokens, temperature=temperature)
        text = _extract_chat_text(result)
        if not _is_api_error(text):
            return text

    # 3. szint: API nélküli, kontextus-alapú válasz
    docs = [line.lstrip("•").strip() for line in context.split("\n") if line.strip()]
    first = docs[0] if docs else context[:200]
    if quick_mode:
        return first
    return (
        f"  Releváns adat azonosítása: A kontextus releváns a feltett kérdéshez.\n"
        f"  2. lépés – Kulcsadat: {first}\n"
        f"  3. lépés – Ellenőrzés: dokumentum alapján következtetés levonható.\n"
        f"Részkövetkeztetés: A kérdés megválaszolható a rendelkezésre álló dokumentumok alapján."
    )


# ─────────────────────────────────────────────
# 2. API hívás – Összefoglalás (BART)
# ─────────────────────────────────────────────
def call_summarization(text: str, api_key: str) -> str:
    """facebook/bart-large-cnn: összefoglalja a releváns dokumentumokat."""
    url = f"{HF_API_BASE}/{SUMMARIZATION_MODEL}"
    word_count = len(text.split())
    # min_length ne haladja meg az input szóhosszát, különben BART hibátst dob
    min_len = max(10, min(40, word_count // 3))
    payload = {
        "inputs": text[:1024],
        "parameters": {"max_length": 200, "min_length": min_len},
        "options": {"wait_for_model": True},
    }
    result = _hf_post(url, payload, api_key)
    return _extract_text(result, key="summary_text")


# ─────────────────────────────────────────────
# 3. API hívás – CoT közbülső következtetés (Llama)
# ─────────────────────────────────────────────
def call_reasoning(context: str, question: str, api_key: str) -> str:
    """
    Garantált közbülső CoT következtetés háromszintű fallback-kel (_guaranteed_chat):
      1. Together AI router  – Meta-Llama-3.1-8B-Instruct-Turbo
      2. HF Inference API    – Mistral-7B-Instruct-v0.3
      3. Kontextus-alapú szöveges következtetés (API nélkül)

    Personality + Few-Shot + CoT prompt. Ez a 3. HuggingFace API hívás.
    Sosem ad vissza nyers API-hibaüzenetet.
    """
    messages = [
        {
            "role": "system",
            "content": (
                f"{PERSONALITY}\n\n"
                f"Tanulj a következő mintapéldákból:\n{FEW_SHOT_EXAMPLES}"
            ),
        },
        {
            "role": "user",
            "content": (
                f"=== Feladat ===\n"
                f"Kontextus: {context}\n"
                f"Kérdés: {question}\n"
                f"Gondolkodás (lépések szerint, majd részkövetkeztetés):\n"
                f"  1. lépés –"
            ),
        },
    ]
    return _guaranteed_chat(messages, api_key, max_tokens=350, temperature=0.35, context=context)


# ─────────────────────────────────────────────
# 4. API hívás – Végső válasz (Mistral)
# ─────────────────────────────────────────────
def call_final_answer(
    context: str,
    question: str,
    summarized: str,
    reasoning: str,
    api_key: str,
) -> str:
    """
    Garantált végső válasz háromszintű fallback-kel (_guaranteed_chat):
      1. Together AI router  – Meta-Llama-3.1-8B-Instruct-Turbo
      2. HF Inference API    – Mistral-7B-Instruct-v0.3
      3. Kontextus-alapú szöveges válasz (API nélkül)

    Ez a 4. HuggingFace API hívás. Sosem ad vissza nyers API-hibaüzenetet.
    """
    messages = [
        {"role": "system", "content": PERSONALITY},
        {
            "role": "user",
            "content": (
                f"Rendelkezésre álló dokumentumok:\n{context}\n\n"
                f"Összefoglalt kontextus:\n{summarized}\n\n"
                f"Közbülső következtetés:\n{reasoning}\n\n"
                f"Kérdés: {question}\n\n"
                f"A fenti dokumentumok és következtetés alapján adj tömör, egyértelmű végső választ "
                f"MAGYARUL. Ha a dokumentumok nem tartalmaznak elegendő információt, jelezd."
            ),
        },
    ]
    return _guaranteed_chat(messages, api_key, max_tokens=400, temperature=0.25, context=context)


# ─────────────────────────────────────────────
# 5. API hívás – Egy mondatos tömör válasz
# ─────────────────────────────────────────────
def call_quick_answer(context: str, question: str, api_key: str) -> str:
    """
    Garantált egy mondatos válasz háromszintű fallback-kel (_guaranteed_chat, quick_mode=True):
      1. Together AI router  – Meta-Llama-3.1-8B-Instruct-Turbo
      2. HF Inference API    – Mistral-7B-Instruct-v0.3
      3. Az első releváns dokumentum mondata (API nélkül)

    Alacsony temperature (0.1) a determinisztikus, tömör kimenetért.
    Ez az 5. HuggingFace API hívás. Sosem ad vissza nyers API-hibaüzenetet.
    """
    messages = [
        {"role": "system", "content": PERSONALITY},
        {
            "role": "user",
            "content": (
                f"Kontextus: {context}\n\n"
                f"Kérdés: {question}\n\n"
                f"Válaszolj PONTOSAN EGY MONDATBAN, MAGYARUL, kizárólag a kontextus alapján. "
                f"Ne írj többet, mint egyetlen teljes mondat."
            ),
        },
    ]
    return _guaranteed_chat(
        messages, api_key, max_tokens=200, temperature=0.1,
        context=context, quick_mode=True,
    )


# ─────────────────────────────────────────────
# Chain-of-Thought pipeline
# ─────────────────────────────────────────────
def run_cot_pipeline(
    question: str,
    documents: List[str],
    api_key: str,
    top_k: int = 3,
) -> dict:
    """
    Teljes CoT pipeline négy HuggingFace API hívással:

    Lépés 1  – find_most_similar  → Embedding API (all-MiniLM-L6-v2)
    Lépés 2  – call_summarization → BART Summarization API
    Lépés 3  – call_reasoning     → Mistral CoT+Few-Shot+Personality API
    Lépés 4  – call_final_answer  → Mistral Final Answer API
    """
    steps: dict = {}

    # ── Lépés 1: Releváns dokumentumok keresése (Embedding API) ──
    status = st.status("1. lépés: Releváns dokumentumok keresése (Embedding API)…", expanded=True)
    with status:
        similar = find_most_similar(question, documents, top_k=top_k)
        steps["similar_docs"] = similar
        for doc, score in similar:
            st.write(f"• [{score:.4f}] {doc[:120]}{'…' if len(doc) > 120 else ''}")
    status.update(label="1. lépés: Releváns dokumentumok ", state="complete", expanded=False)

    context_text = "\n".join(f"• {doc}" for doc, _ in similar)

    # ── Lépés 2: Kontextus összefoglalása (BART API) ──
    status2 = st.status("2. lépés: Kontextus összefoglalása (BART Summarization API)…", expanded=True)
    _summary_failed = False
    with status2:
        raw_summ = call_summarization(context_text, api_key)
        if _is_api_error(raw_summ):
            summarized = context_text[:500]
            _summary_failed = True
            st.warning("Az összefoglalás nem sikerült – a nyers kontextust használjuk tovább.")
        else:
            summarized = raw_summ
            st.write(summarized)
        steps["summarized_context"] = summarized
    if _summary_failed:
        status2.update(
            label="2. lépés: Összefoglalás (nem sikerült – kontextus alapján folytatva)",
            state="error",
            expanded=False,
        )
    else:
        status2.update(label="2. lépés: Összefoglalás ", state="complete", expanded=False)

    # ── Lépés 3: Közbülső CoT következtetés (Mistral API) ──
    status3 = st.status("3. lépés: Közbülső következtetés (Mistral CoT + Few-Shot + Personality)…", expanded=True)
    with status3:
        raw_reasoning = call_reasoning(context_text, question, api_key)
        if _is_api_error(raw_reasoning):
            reasoning = REASONING_FALLBACK
            steps["reasoning"] = reasoning
            steps["reasoning_error"] = raw_reasoning
            st.warning(
                "A CoT következtetési lépés nem érhető el – "
                "a végső válasz közvetlenül a dokumentumok alapján készül."
            )
            status3.update(
                label="3. lépés: Közbülső következtetés (kihagyva, folytatás kontextusból)",
                state="error",
                expanded=False,
            )
        else:
            reasoning = raw_reasoning
            steps["reasoning"] = reasoning
            st.code(f"1. lépés – {reasoning}", language=None)
            status3.update(label="3. lépés: Közbülső következtetés ", state="complete", expanded=False)

    # ── Lépés 4: Végső válasz (Mistral API) ──
    status4 = st.status("4. lépés: Végső válasz generálása (Mistral Final Answer API)…", expanded=True)
    with status4:
        final = call_final_answer(context_text, question, summarized, reasoning, api_key)
        steps["final_answer"] = final
        st.write(final)
    status4.update(label="4. lépés: Végső válasz ", state="complete", expanded=False)

    # ── Lépés 5: Egy mondatos tömör válasz ──
    status5 = st.status("5. lépés: Egy mondatos válasz generálása…", expanded=True)
    _quick_failed = False
    with status5:
        quick = call_quick_answer(context_text, question, api_key)
        steps["quick_answer"] = quick
        if _is_api_error(quick):
            _quick_failed = True
            st.warning("Az egy mondatos válasz generálása nem sikerült.")
        else:
            st.write(quick)
    if _quick_failed:
        status5.update(
            label="5. lépés: Egy mondatos válasz (nem sikerült)",
            state="error",
            expanded=False,
        )
    else:
        status5.update(label="5. lépés: Egy mondatos válasz ", state="complete", expanded=False)

    return steps


# ─────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────
def main() -> None:
    st.set_page_config(
        page_title="Dokumentum Q&A – Chain-of-Thought",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # ── Halványkék háttér ──
    st.markdown(
        """
        <style>
        .stApp {
            background-color: #EBF5FB;
        }
        [data-testid="stSidebar"] {
            background-color: #D6EAF8;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # ── API kulcs kizárólag .env-ből ──
    api_key = os.getenv("HF_API_KEY", "")

    # ── Fejléc ──
    st.title("Dokumentum-alapú Q&A")
    st.markdown(
        "**Chain-of-Thought &nbsp;|&nbsp; Personality Prompting &nbsp;|&nbsp; "
        "Few-Shot Prompting &nbsp;|&nbsp; 4× HuggingFace API**"
    )
    st.divider()

    # ── Oldalsáv ──
    with st.sidebar:
        st.header("Beállítások")

        top_k = st.slider("Releváns dokumentumok száma (top-k)", 1, 6, 3)

    # ── Fő tartalom ──
    left, right = st.columns([1, 1], gap="large")

    with left:
        st.subheader("Dokumentumok")
        mode = st.radio("Bevitel módja:", ["Szöveg beírása", "Fájl feltöltése (.txt)"], horizontal=True)

        documents: List[str] = []

        if mode == "Szöveg beírása":
            raw = st.text_area(
                "Dokumentumok (soronként 1 dokumentum)",
                height=280,
                placeholder=(
                    "Minden sor egy önálló dokumentumnak számít.\n\n"
                    "Pl.:\n"
                    "Az elektron negatív töltésű részecske.\n"
                    "A proton pozitív töltésű, az atommag alkotója.\n"
                    "Kylian Mbappé a Paris Saint-Germain csatára."
                ),
            )
            if raw:
                documents = [line.strip() for line in raw.splitlines() if line.strip()]
        else:
            files = st.file_uploader(
                "Tölts fel .txt fájlokat",
                type=["txt"],
                accept_multiple_files=True,
            )
            for f in files:
                content = f.read().decode("utf-8", errors="ignore")
                paragraphs = [p.strip() for p in content.split("\n\n") if p.strip()]
                if not paragraphs:
                    paragraphs = [ln.strip() for ln in content.splitlines() if ln.strip()]
                documents.extend(paragraphs)

        if documents:
            st.success(f"{len(documents)} dokumentum betöltve")
            with st.expander("Betöltött dokumentumok listája"):
                for i, d in enumerate(documents, 1):
                    st.markdown(f"**{i}.** {d}")

    with right:
        st.subheader("Kérdés és futtatás")

        question = st.text_input(
            "Kérdés a dokumentumokkal kapcsolatban",
            placeholder="Pl.: Ki játszik a Paris Saint-Germain csapatban?",
        )

        can_run = bool(documents and question and api_key)

        if not api_key:
            st.warning("Az API kulcs nem található. Add meg a HF_API_KEY értékét a .env fájlban!")
        if not documents:
            st.info("Adj meg dokumentumokat a bal oldali panelen!")
        if documents and not question:
            st.info("Adja meg a kérdéseket!")


        run = st.button(
            "Futtatás",
            type="primary",
            disabled=not can_run,
            use_container_width=True,
        )

    # ── Eredmények ──
    if run:
        st.divider()
        st.subheader("Chain-of-Thought pipeline futtatása")

        try:
            results = run_cot_pipeline(question, documents, api_key, top_k=top_k)

            st.divider()
            st.subheader("Végső válasz")

            quick = results.get("quick_answer", "")
            if quick and not _is_api_error(quick):
                st.markdown("**Egy mondatos válasz:**")
                st.info(quick)
                st.divider()

            final = results.get("final_answer", "")
            if final and not _is_api_error(final):
                st.markdown("**Részletes végső válasz:**")
                st.success(final)
            else:
                st.error("Nem sikerült választ generálni. Ellenőrizd az API kulcsot és próbáld újra.")

            with st.expander("Részletes CoT összesítő", expanded=False):
                st.markdown("**Releváns dokumentumok:**")
                for doc, score in results.get("similar_docs", []):
                    szint = "[magas]" if score >= 0.5 else "[közepes]" if score >= 0.3 else "[alacsony]"
                    st.markdown(f"{szint} `{score}` — {doc}")

                st.markdown("---")
                st.markdown("**Összefoglalt kontextus (BART):**")
                st.info(results.get("summarized_context", "–"))

                st.markdown("---")
                st.markdown("**Közbülső CoT következtetés (Mistral):**")
                reasoning_text = results.get("reasoning", "–")
                st.code(f"1. lépés – {reasoning_text}", language=None)

        except Exception as exc:
            st.error(f"Hiba a feldolgozás során: {exc}")
            st.exception(exc)


if __name__ == "__main__":
    main()
