// Nawi Saadi brand lockup — shared by the app header and the login page.
import { useEffect, useState } from "react";

const GOLD = "#c9a23a";

/** The Nawi Saadi gold arch / kufic dome mark, recreated as crisp vector. */
export function BrandMark({ className = "" }: { className?: string }) {
  // Graduated minaret bars rising to a central apex, under a double arch.
  const bars = [24, 31.2, 38.4, 45.6, 52.8, 60, 67.2, 74.4, 81.6, 88.8, 96].map((x) => {
    const h = 14 + 26 * (1 - Math.abs(x - 60) / 36);
    return { x, top: 60 - h };
  });
  return (
    <svg viewBox="0 0 120 70" className={className} aria-hidden>
      <path
        d="M8 62 C8 26 60 9 60 9 C60 9 112 26 112 62"
        fill="none"
        stroke={GOLD}
        strokeWidth="3.4"
        strokeLinecap="round"
      />
      <path
        d="M16 62 C16 33 60 18 60 18 C60 18 104 33 104 62"
        fill="none"
        stroke={GOLD}
        strokeWidth="1.1"
        opacity="0.5"
      />
      <g stroke={GOLD} strokeWidth="2.9" strokeLinecap="round">
        {bars.map((b, i) => (
          <line key={i} x1={b.x} y1={60} x2={b.x} y2={b.top} />
        ))}
      </g>
    </svg>
  );
}

/** Vector lockup of the full brand: arch mark + wordmark + Arabic subtitle. */
export function BrandLogoVector({ light = true }: { light?: boolean }) {
  return (
    <div className="flex items-center gap-3">
      <BrandMark className="h-11 w-auto shrink-0" />
      <div className="leading-tight">
        <div
          className={`text-[17px] font-semibold tracking-[0.26em] ${light ? "text-white" : "text-slate-800"}`}
        >
          NAWI SAADI
        </div>
        <div className="flex items-center gap-2">
          <span
            className="text-[8.5px] font-semibold tracking-[0.32em]"
            style={{ color: GOLD }}
          >
            TRAVEL &amp; TOURISM
          </span>
          <span className="text-[9px]" style={{ color: GOLD }} dir="rtl">
            للسفر والسياحة
          </span>
        </div>
      </div>
    </div>
  );
}

/**
 * Company logo. Shows the polished vector brand immediately, and silently
 * upgrades to your exact logo image if it is present at:
 *   public/nawi-saadi-logo.png
 * (Preloaded so a missing file never flashes a broken-image icon.)
 */
export function BrandLogo({ light = true }: { light?: boolean }) {
  const [hasImg, setHasImg] = useState(false);
  useEffect(() => {
    const img = new Image();
    img.onload = () => {
      if (img.naturalWidth > 1 && img.naturalHeight > 1) setHasImg(true);
    };
    img.src = "/nawi-saadi-logo.png";
  }, []);
  if (hasImg)
    return (
      <img
        src="/nawi-saadi-logo.png"
        alt="Nawi Saadi Travel & Tourism"
        className="h-12 w-auto shrink-0 object-contain"
        style={{ maxWidth: 240 }}
      />
    );
  return <BrandLogoVector light={light} />;
}
