# routing/web_scorer.py
"""
Web content scoring system for determining if a prompt needs web search.
"""
import re
import os
from typing import List, Set, Optional
from pathlib import Path
from config.settings_simple import get_settings


class WebScorer:
    """Scores prompts to determine if they need web search capabilities."""
    
    def __init__(self):
        self.settings = get_settings()
        self._load_web_terms()
        self._compile_patterns()
    
    def _load_web_terms(self):
        """Load web terms from file or use defaults."""
        terms_path = Path("web_terms.txt")
        if terms_path.exists():
            try:
                with open(terms_path, "r", encoding="utf-8") as f:
                    self.web_terms = tuple(
                        t.strip().lower()
                        for t in f
                        if t.strip() and not t.strip().startswith("#")
                    )
            except Exception:
                self.web_terms = self._get_default_web_terms()
        else:
            self.web_terms = self._get_default_web_terms()
    
    def _get_default_web_terms(self) -> tuple:
        """Get default web terms for scoring."""
        return (
            # Recency/news
            "latest", "today", "current", "currently", "right now", "present",
            "real-time", "breaking", "live", "headline", "just announced",
            "recent", "updated", "update", "outlook", "forecast", "projection",
            "guidance", "trends", "trend", "news", "article", "report",
            "press release", "coverage", "announcement", "commentary", "bulletin",
            
            # Macro/policy/prints
            "interest rate", "interest rates", "rates", "fed", "federal reserve",
            "fomc", "dot plot", "inflation", "cpi", "ppi", "pce", "jobs report",
            "nonfarm payrolls", "unemployment", "ism", "gdp", "retail sales",
            "housing starts", "consumer confidence", "beige book", "ecb", "boe",
            "boj", "rba", "snb", "banxico", "central bank", "policy rate",
            
            # Markets
            "bond yield", "yield", "yields", "treasury", "ust", "2y", "10y", "30y",
            "curve", "sofr", "libor", "euribor", "bps", "bp", "stock", "equity",
            "index", "s&p", "spx", "sp500", "dow", "nasdaq", "ndx", "russell",
            "volatility", "vix", "option flow", "open interest",
            
            # Commodities/energy
            "oil", "crude", "brent", "wti", "nat gas", "natural gas", "lng",
            "uranium", "gold", "silver", "copper", "lithium", "nickel", "coal",
            "power prices", "spark spread",
            
            # FX/crypto
            "fx", "forex", "usd", "eur", "jpy", "gbp", "cny", "aud", "cad", "chf",
            "mxn", "brl", "dxy", "us dollar index", "euro", "yen", "pound",
            "crypto", "bitcoin", "btc", "ethereum", "eth", "solana", "sol",
            "etf approval",
            
            # Corporate actions / filings
            "m&a", "deal", "transaction", "takeover", "bid", "offer", "ipo",
            "spac", "lbo", "private equity", "venture capital", "fundraise",
            "bond issue", "debt offering", "tender", "consent solicitation",
            "sec", "edgar", "filing", "prospectus", "registration statement",
            "amendment", "proxy", "10-k", "10k", "10-q", "10q", "8-k", "6-k", "20-f",
            
            # Sector news
            "earnings", "eps", "revenue", "pre-announcement", "data center",
            "ai chips", "semiconductor", "foundry", "airline capacity",
            "load factor", "hotel revpar", "same-store sales", "ports congestion",
            "freight rates", "container rates",
            
            # Other fast domains
            "weather", "hurricane", "storm path", "earthquake", "election",
            "poll", "ballot", "turnout", "sports score", "scoreboard", "odds",
            "spread", "release notes", "changelog", "version", "zero-day", "cve",
            
            # Web cues
            "google", "search", "web", "internet", "look up", "find", "browse",
        )
    
    def _compile_patterns(self):
        """Compile regex patterns for web scoring."""
        # Date patterns
        months = r"(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)"
        self.date_patterns = [
            re.compile(rf"\b{months}\.?[-/ ]?\d{{1,2}}(,? ?\d{{2,4}})?\b", re.I),
            re.compile(r"\b\d{4}-\d{2}-\d{2}\b"),
            re.compile(r"\bq[1-4]\s*[-/]?\s*\d{2,4}\b", re.I),
            re.compile(r"\b(ytd|mtd|qoq|yoy)\b", re.I),
        ]
        
        # Time words
        self.time_words = re.compile(
            r"\b(today|tonight|this (week|month|quarter|year)|yesterday|tomorrow|now|as of)\b",
            re.I,
        )
        
        # Percent or basis points
        self.percent_or_bps = re.compile(r"\b\d+(\.\d+)?\s*(%|percent|bps|bp)\b", re.I)
        
        # Price symbols
        self.price_syms = re.compile(r"(\$|€|£|¥)\s?\d{1,3}(,\d{3})*(\.\d+)?")
        
        # Ticker symbols
        self.ticker = re.compile(r"\b[A-Z]{1,5}(\.[A-Z])?\b")
        
        # Common caps to exclude from ticker scoring
        self.common_caps = {"A", "I", "AND", "THE", "USD", "CEO", "CFO", "EPS", "Q", "BP", "AI"}
    
    def score_webbiness(self, text: str) -> int:
        """Score how webby a text is (higher = more webby)."""
        if not text:
            return 0
        
        s = text.lower()
        score = 0
        
        # Keywords
        score += sum(1 for w in self.web_terms if w in s)
        
        # Date patterns
        score += sum(1 for rx in self.date_patterns if rx.search(text))
        
        # Time words
        if self.time_words.search(text):
            score += 1
        
        # Percent or basis points
        if self.percent_or_bps.search(text):
            score += 1
        
        # Price symbols
        if self.price_syms.search(text):
            score += 1
        
        # Ticker density (avoid common caps)
        tickers = [
            m.group(0) for m in self.ticker.finditer(text)
            if m.group(0) not in self.common_caps
        ]
        if len(tickers) >= 2:
            score += 1
        if len(tickers) >= 4:
            score += 1
        
        return score
    
    def looks_webby(self, text: str) -> bool:
        """Determine if a prompt looks like it needs web search."""
        try:
            return self.score_webbiness(text) >= self.settings.web_scoring_threshold
        except Exception:
            return False
    
    def get_web_score(self, text: str) -> int:
        """Get the raw web score for a text."""
        return self.score_webbiness(text)


# Global web scorer instance
_web_scorer: Optional[WebScorer] = None


def get_web_scorer() -> WebScorer:
    """Get the global web scorer instance."""
    global _web_scorer
    if _web_scorer is None:
        _web_scorer = WebScorer()
    return _web_scorer
