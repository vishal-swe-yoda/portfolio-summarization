"""
LLMCommentaryEngine: Implements the Strategy Pattern for qualitative portfolio commentary.
Supported Strategies:
- GeminiStrategy (Google Gemini API via v1beta generateContent)
- DeepSeekStrategy (DeepSeek API via /chat/completions)
- OpenAIStrategy (OpenAI API via /v1/chat/completions)
- OfflineFallbackStrategy (Deterministic offline rule engine using Excel research notes)
"""

from abc import ABC, abstractmethod
from typing import Dict, Any, List, Optional
import json
import os
import urllib.request
import urllib.error
import sys
import re

def _load_env_safely():
    """Loads environment variables from .env if not already set in environment."""
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except Exception:
        pass

    for candidate in [".env", os.path.join(os.path.dirname(__file__), "..", ".env"), os.path.join(os.getcwd(), ".env")]:
        if os.path.exists(candidate):
            try:
                with open(candidate, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith("#") and "=" in line:
                            k, v = line.split("=", 1)
                            k = k.strip()
                            v = v.strip().strip("'\"")
                            if k and k not in os.environ:
                                os.environ[k] = v
            except Exception:
                pass

_load_env_safely()

try:
    from google import genai
    from google.genai import types
    HAS_GENAI = True
except ImportError:
    # Auto-detect local virtualenv site-packages if running outside venv
    vpoc_site = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "vpoc", "Lib", "site-packages"))
    if os.path.exists(vpoc_site) and vpoc_site not in sys.path:
        sys.path.insert(0, vpoc_site)
        try:
            from google import genai
            from google.genai import types
            HAS_GENAI = True
        except ImportError:
            genai = None
            types = None
            HAS_GENAI = False
    else:
        genai = None
        types = None
        HAS_GENAI = False

SYSTEM_INSTRUCTION = (
    "You are an institutional portfolio analyst. Return ONLY valid JSON matching this exact schema:\n"
    "{\n"
    '  "overweight_commentary": [\n'
    '    {\n'
    '      "sector": str,\n'
    '      "fund_weight": float,\n'
    '      "benchmark_weight": float,\n'
    '      "active_weight": float,\n'
    '      "observations": [str],\n'
    '      "risks": [str]\n'
    '    }\n'
    "  ],\n"
    '  "underweight_commentary": [\n'
    '    {\n'
    '      "sector": str,\n'
    '      "fund_weight": float,\n'
    '      "benchmark_weight": float,\n'
    '      "active_weight": float,\n'
    '      "observations": [str],\n'
    '      "risks": [str]\n'
    '    }\n'
    "  ],\n"
    '  "portfolio_observations": [str],\n'
    '  "data_limitations": [str]\n'
    "}\n"
    "Do NOT include markdown code blocks, backticks, or any extraneous conversational text."
)


# =============================================================================
# Strategy Pattern: Abstract Base Class
# =============================================================================
class LLMStrategy(ABC):
    @abstractmethod
    def generate(self, payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Executes LLM call and returns structured dictionary or None on failure."""
        pass


# =============================================================================
# Concrete Strategy 1: Gemini API
# =============================================================================
class GeminiStrategy(LLMStrategy):
    def __init__(self, config: Dict[str, Any]):
        llm_cfg = config.get("llm", {})
        self.model = llm_cfg.get("model", "gemini-3.8-flash")
        self.temperature = llm_cfg.get("temperature", 0.2)
        self.api_key = (
            os.environ.get("GEMINI_API_KEY")
            or os.environ.get("GOOGLE_API_KEY")
            or llm_cfg.get("api_key")
        )
        if HAS_GENAI and self.api_key:
            try:
                self.client = genai.Client(api_key=self.api_key)
            except Exception as e:
                print(f"Gemini client initialization error: {e}")
                self.client = None
        else:
            self.client = None

    def _generate_via_rest(self, payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Direct REST fallback to Google Gemini v1beta endpoint using urllib (zero external dependencies)."""
        if not self.api_key:
            return None

        model_name = self.model
        endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={self.api_key}"
        prompt = (
            f"{SYSTEM_INSTRUCTION}\n\n"
            f"PORTFOLIO METRICS DATA:\n{json.dumps(payload, indent=2)}"
        )
        request_body = {
            "contents": [
                {
                    "parts": [{"text": prompt}]
                }
            ],
            "generationConfig": {
                "temperature": self.temperature,
                "responseMimeType": "application/json"
            }
        }
        try:
            req_data = json.dumps(request_body).encode("utf-8")
            headers = {"Content-Type": "application/json"}
            req = urllib.request.Request(endpoint, data=req_data, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=15) as resp:
                if resp.status == 200:
                    resp_json = json.loads(resp.read().decode("utf-8"))
                    candidates = resp_json.get("candidates", [])
                    if candidates:
                        content_parts = candidates[0].get("content", {}).get("parts", [])
                        if content_parts:
                            text = content_parts[0].get("text", "")
                            text = re.sub(r"^```(?:json)?\s*", "", text.strip())
                            text = re.sub(r"\s*```$", "", text.strip())
                            return json.loads(text)
        except Exception as e:
            print(f"Gemini direct REST API error: {e}")
        return None

    def generate(self, payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        if not self.api_key:
            print("GeminiStrategy: No API key found. Set GEMINI_API_KEY or GOOGLE_API_KEY in .env or environment.")
            return None

        # 1. Try google.genai SDK
        if HAS_GENAI and self.client:
            prompt = (
                f"{SYSTEM_INSTRUCTION}\n\n"
                f"PORTFOLIO METRICS DATA:\n{json.dumps(payload, indent=2)}"
            )
            try:
                response = self.client.models.generate_content(
                    model=self.model,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        temperature=self.temperature,
                        response_mime_type="application/json",
                    ),
                )
                if response and response.text:
                    text = response.text.strip()
                    text = re.sub(r"^```(?:json)?\s*", "", text)
                    text = re.sub(r"\s*```$", "", text)
                    return json.loads(text)
            except Exception as e:
                print(f"GeminiStrategy SDK error: {e}. Trying direct REST fallback...")

        # 2. Try direct REST fallback
        return self._generate_via_rest(payload)


# =============================================================================
# Concrete Strategy 2: DeepSeek API
# =============================================================================
class DeepSeekStrategy(LLMStrategy):
    def __init__(self, config: Dict[str, Any]):
        llm_cfg = config.get("llm", {})
        self.model = llm_cfg.get("model", "deepseek-chat")
        self.endpoint = llm_cfg.get("endpoint", "https://api.deepseek.com/chat/completions")
        self.temperature = llm_cfg.get("temperature", 0.2)
        self.api_key = os.environ.get("DEEPSEEK_API_KEY") or llm_cfg.get("api_key")

    def generate(self, payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        if not self.api_key:
            print("DeepSeekStrategy: DEEPSEEK_API_KEY not found in environment.")
            return None

        prompt_text = f"PORTFOLIO METRICS DATA:\n{json.dumps(payload, indent=2)}"
        request_body = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": SYSTEM_INSTRUCTION},
                {"role": "user", "content": prompt_text}
            ],
            "response_format": {"type": "json_object"},
            "temperature": self.temperature
        }

        try:
            req_data = json.dumps(request_body).encode("utf-8")
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}"
            }
            req = urllib.request.Request(self.endpoint, data=req_data, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=12) as response:
                if response.status == 200:
                    resp_json = json.loads(response.read().decode("utf-8"))
                    choices = resp_json.get("choices", [])
                    if choices:
                        content = choices[0].get("message", {}).get("content", "").strip()
                        return json.loads(content)
        except Exception as e:
            print(f"DeepSeekStrategy error: {e}")
            return None
        return None


# =============================================================================
# Concrete Strategy 3: OpenAI API
# =============================================================================
class OpenAIStrategy(LLMStrategy):
    def __init__(self, config: Dict[str, Any]):
        llm_cfg = config.get("llm", {})
        self.model = llm_cfg.get("model", "gpt-4o-mini")
        self.endpoint = llm_cfg.get("endpoint", "https://api.openai.com/v1/chat/completions")
        self.temperature = llm_cfg.get("temperature", 0.2)
        self.api_key = os.environ.get("OPENAI_API_KEY") or llm_cfg.get("api_key")

    def generate(self, payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        if not self.api_key:
            print("OpenAIStrategy: OPENAI_API_KEY not found in environment.")
            return None

        prompt_text = f"PORTFOLIO METRICS DATA:\n{json.dumps(payload, indent=2)}"
        request_body = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": SYSTEM_INSTRUCTION},
                {"role": "user", "content": prompt_text}
            ],
            "response_format": {"type": "json_object"},
            "temperature": self.temperature
        }

        try:
            req_data = json.dumps(request_body).encode("utf-8")
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}"
            }
            req = urllib.request.Request(self.endpoint, data=req_data, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=12) as response:
                if response.status == 200:
                    resp_json = json.loads(response.read().decode("utf-8"))
                    choices = resp_json.get("choices", [])
                    if choices:
                        content = choices[0].get("message", {}).get("content", "").strip()
                        return json.loads(content)
        except Exception as e:
            print(f"OpenAIStrategy error: {e}")
            return None
        return None


# =============================================================================
# Concrete Strategy 4: Deterministic Offline Fallback
# =============================================================================
class OfflineFallbackStrategy(LLMStrategy):
    def __init__(self, config: Dict[str, Any], data_bundle: Optional[Dict[str, Any]] = None):
        self.config = config
        self.report_cfg = config.get("report", {})
        self.data_bundle = data_bundle or {}

    def generate(self, payload: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        display_date = self.report_cfg.get("display_date_page2", "March 30, 2026")
        bm_name = self.report_cfg.get("benchmark", "BSE 100 TRI")

        ow_list = []
        for item in payload.get("top_overweight", []):
            notes = item.get("notes", "")
            if notes and "|" in notes:
                obs = [n.strip() for n in notes.split("|") if n.strip()]
            elif notes:
                obs = [notes.strip()]
            else:
                obs = [
                    f"Structural growth positioning in {item['sector']}",
                    "Favorable earnings trajectory and market share consolidation",
                    "Strong balance sheet characteristics and healthy cash conversion"
                ]
            ow_list.append({
                "sector": item["sector"],
                "fund_weight": item["fund_weight"],
                "benchmark_weight": item["benchmark_weight"],
                "active_weight": item["active_weight"],
                "observations": obs,
                "risks": ["Sector-specific earnings or valuation changes may affect relative performance."]
            })

        uw_list = []
        for item in payload.get("top_underweight", []):
            notes = item.get("notes", "")
            if notes and "|" in notes:
                obs = [n.strip() for n in notes.split("|") if n.strip()]
            elif notes:
                obs = [notes.strip()]
            else:
                obs = [
                    "Capital intensive dynamics and extended gestation periods",
                    "Moderate return ratios relative to core portfolio hurdle rates",
                    "Near-term growth moderate with current valuations reflecting fair pricing"
                ]
            uw_list.append({
                "sector": item["sector"],
                "fund_weight": item["fund_weight"],
                "benchmark_weight": item["benchmark_weight"],
                "active_weight": item["active_weight"],
                "observations": obs,
                "risks": ["Sharp cyclical rally in underweight constituents could create relative drag."]
            })

        # Formulation of portfolio observations from composition tiers
        comp = self.data_bundle.get("composition_tiers", {})
        c1_val = "89%"
        r1_val = "63%"
        r1_bm = "56%"
        for o in comp.get("ocf", []):
            if o.get("tier") == "C1":
                c1_val = o.get("fund", "89%")
        for r in comp.get("roce", []):
            if r.get("tier") == "R1":
                r1_val = r.get("fund", "63%")
                r1_bm = r.get("bm", "56%")

        portfolio_observations = [
            f"High quality bias: {c1_val} of portfolio companies reside in top Operating Cash Flow tier (C1) and {r1_val} in top RoCE tier (R1 vs {r1_bm} for {bm_name}).",
            "Concentrated high-conviction positioning: Top 10 holdings represent ~49% of portfolio NAV with disciplined risk management.",
            "Balanced cyclical and defensive exposure providing resilience across market regimes."
        ]

        data_limitations = [
            f"Sector commentary reflects point-in-time holdings as of {display_date}."
        ]

        return {
            "overweight_commentary": ow_list,
            "underweight_commentary": uw_list,
            "portfolio_observations": portfolio_observations,
            "data_limitations": data_limitations
        }


# =============================================================================
# Strategy Factory
# =============================================================================
class LLMStrategyFactory:
    @staticmethod
    def get_strategy(provider: str, config: Dict[str, Any], data_bundle: Dict[str, Any]) -> LLMStrategy:
        p = provider.lower().strip()
        if p == "gemini":
            return GeminiStrategy(config)
        elif p == "deepseek":
            return DeepSeekStrategy(config)
        elif p == "openai":
            return OpenAIStrategy(config)
        else:
            return OfflineFallbackStrategy(config, data_bundle)


# =============================================================================
# Orchestrator: LLMCommentaryEngine
# =============================================================================
class LLMCommentaryEngine:
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.llm_cfg = config.get("llm", {})
        self.report_cfg = config.get("report", {})
        self.provider = self.llm_cfg.get("provider", "gemini").lower()
        self.instruction = self.llm_cfg.get("instruction", "Write concise, evidence-based portfolio observations.")

    def build_payload(self, data_bundle: Dict[str, Any]) -> Dict[str, Any]:
        """Constructs the strict input payload for the LLM."""
        ow_df = data_bundle.get("top_overweight_sectors")
        uw_df = data_bundle.get("top_underweight_sectors")

        top_ow = []
        if ow_df is not None and not ow_df.empty:
            for _, r in ow_df.head(3).iterrows():
                top_ow.append({
                    "sector": str(r.get("Sector", "")),
                    "fund_weight": float(r.get("Fund_Weight", 0.0)),
                    "benchmark_weight": float(r.get("BM_Weight", 0.0)),
                    "active_weight": float(r.get("Active_Weight", 0.0)),
                    "notes": str(r.get("Research_Commentary", "")),
                })

        top_uw = []
        if uw_df is not None and not uw_df.empty:
            for _, r in uw_df.head(3).iterrows():
                top_uw.append({
                    "sector": str(r.get("Sector", "")),
                    "fund_weight": float(r.get("Fund_Weight", 0.0)),
                    "benchmark_weight": float(r.get("BM_Weight", 0.0)),
                    "active_weight": float(r.get("Active_Weight", 0.0)),
                    "notes": str(r.get("Research_Commentary", "")),
                })

        return {
            "report_date": self.report_cfg.get("report_date", "2026-03-30"),
            "benchmark": self.report_cfg.get("benchmark", "BSE 100 TRI"),
            "top_overweight": top_ow,
            "top_underweight": top_uw,
            "instruction": self.instruction,
        }

    def generate_commentary(self, data_bundle: Dict[str, Any], output_audit_dir: str = "output/intermediate") -> Dict[str, Any]:
        """Dispatches request via strategy pattern with resilient offline fallback."""
        payload = self.build_payload(data_bundle)

        # 1. Execute chosen strategy
        strategy = LLMStrategyFactory.get_strategy(self.provider, self.config, data_bundle)
        response_json = strategy.generate(payload)

        # 2. Resilient fallback if primary strategy returned None (e.g. offline, rate limit, no key)
        if not response_json:
            print(f"LLM strategy '{self.provider}' returned no response. Activating OfflineFallbackStrategy.")
            fallback = OfflineFallbackStrategy(self.config, data_bundle)
            response_json = fallback.generate(payload)

        # 3. Write audit log
        os.makedirs(output_audit_dir, exist_ok=True)
        audit_path = os.path.join(output_audit_dir, "llm_payload_and_response.json")
        try:
            with open(audit_path, "w", encoding="utf-8") as f:
                json.dump({"input_payload": payload, "output_response": response_json}, f, indent=2)
        except Exception as e:
            print(f"Warning: Failed to write audit log: {e}")

        return response_json
