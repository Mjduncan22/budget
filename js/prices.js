// Live prices. Crypto: CoinGecko (allows browser requests). Funds/stocks: the "Last Price" the
// sync saves from Fidelity in the Holdings tab (stock quote sites block browser requests).
export const DOLLAR_SYMBOLS = new Set(["USD", "SPAXX", "FDRXX", "FZFXX", "CORE", "USDC"]);

/** Ticker -> CoinGecko id. Add a line here if you buy a new coin. */
export const COINGECKO_IDS = {
  BTC: "bitcoin", ETH: "ethereum", DOT: "polkadot", LINK: "chainlink", SOL: "solana", RENDER: "render-token",
  ONDO: "ondo-finance", AERO: "aerodrome-finance", FET: "fetch-ai", JASMY: "jasmycoin", KAS: "kaspa",
  AVAX: "avalanche-2", ADA: "cardano", SUI: "sui", MORPHO: "morpho", PLUME: "plume", CTSI: "cartesi",
  XRP: "ripple", DOGE: "dogecoin", MATIC: "matic-network", POL: "polygon-ecosystem-token", HBAR: "hedera-hashgraph",
  NEAR: "near", LTC: "litecoin", SHIB: "shiba-inu", ATOM: "cosmos", UNI: "uniswap", XLM: "stellar",
};

export async function quotes(holdings) {
  const out = {};
  const crypto = new Set();
  for (const h of holdings) {
    if (DOLLAR_SYMBOLS.has(h.symbol) || h.assetType.toLowerCase() === "cash") out[h.symbol] = { price: 1, change: 0 };
    else if (h.assetType.toLowerCase() === "crypto") crypto.add(h.symbol);
    else if (h.lastPrice) out[h.symbol] = { price: h.lastPrice, change: null };
  }
  const ids = [...crypto].filter((s) => COINGECKO_IDS[s]);
  if (ids.length) {
    const url = `https://api.coingecko.com/api/v3/simple/price?ids=${ids.map((s) => COINGECKO_IDS[s]).join(",")}&vs_currencies=usd&include_24hr_change=true`;
    const r = await fetch(url);
    if (r.ok) {
      const j = await r.json();
      for (const s of ids) {
        const o = j[COINGECKO_IDS[s]];
        if (o && typeof o.usd === "number") out[s] = { price: o.usd, change: o.usd_24h_change ?? null };
      }
    }
  }
  return out;
}
