import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { initializeApp } from "firebase/app";
import { getAuth, onAuthStateChanged, signOut,
  createUserWithEmailAndPassword, signInWithEmailAndPassword,
  sendPasswordResetEmail, signInWithPopup,
  GoogleAuthProvider, OAuthProvider } from "firebase/auth";
import { getFirestore, doc, getDoc, setDoc, deleteDoc } from "firebase/firestore";

/* Firebase config */
const firebaseApp = initializeApp({
  apiKey: "AIzaSyCZTrJTGH8Jh5WBMhMrV39mjKddRj7p78w",
  authDomain: "zafi-524b8.firebaseapp.com",
  projectId: "zafi-524b8",
  storageBucket: "zafi-524b8.firebasestorage.app",
  messagingSenderId: "308516673564",
  appId: "1:308516673564:web:9410954d5fc50fd56667d9"
});
const auth = getAuth(firebaseApp);
const db = getFirestore(firebaseApp);

/* =========================================================================
   ZAFI — finanzas personales con IA
   - Onboarding conversacional (chatbot tipo Claude) que arma categorías + cuentas
   - Captura de ingresos / egresos
   - Auto-categorización: palabras clave locales -> Claude API -> te pregunta
   - Aprende de tus correcciones (guarda palabras clave)
   - Persiste con window.storage (sobrevive entre sesiones)
   ========================================================================= */

/* ----------------------------- estilos ---------------------------------- */
const STYLE = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600;9..144,700&family=Montserrat:wght@300;400;500;600;700&display=swap');

html,body{min-height:100vh;}
body{
  background-color:var(--bg)!important;
}
/* video background - see <video> in JSX */
.cc-video-bg{position:fixed;inset:0;z-index:-1;overflow:hidden;}
.cc-video-bg video{width:100%;height:100%;object-fit:cover;
  filter:blur(10px);transform:scale(1.06);}
#root{background:transparent!important;min-height:100vh;}
.cc-root *{box-sizing:border-box;margin:0;padding:0;}
:root{
  /* Paleta Zafi — glassmorphism plata/azul frío */
  --bg:#DCE1E8;
  --bg-2:#D2D8E1;
  --paper:rgba(255,255,255,.12);
  --paper-solid:#FFFFFF;
  --surface:rgba(255,255,255,.14);
  --surface-2:rgba(255,255,255,.1);
  --surface-3:rgba(20,30,45,.05);
  --ink:#1B2230;
  --bar-fill:#7E8AA0;
  --ink-soft:#6B7585;
  --ink-faint:#8B95A6;
  --line:rgba(30,40,60,.09);
  --line-soft:rgba(30,40,60,.045);
  --green:#1A7A6E;
  --green-2:#22917F;
  --green-soft:rgba(26,122,110,.10);
  --green-glow:rgba(26,122,110,.20);
  --coral:#B5453A;
  --coral-2:#CC5548;
  --coral-soft:rgba(181,69,58,.08);
  --coral-glow:rgba(181,69,58,.12);
  --gold:#8C7FAE;
  --gold-soft:rgba(140,127,174,.10);
  --gold-glow:rgba(140,127,174,.18);

  /* Acentos del orb IA — iridiscente */
  --orb-purple:#A78BFA;
  --orb-blue:#60A5FA;
  --orb-mint:#5EEAD4;

  /* Degradado de acento (barras de categoría, cuenta seleccionada) */
  --accent-grad:linear-gradient(90deg, #0F2A4A 0%, #1D4ED8 45%, #38BDF8 100%);
  --accent-grad-soft:linear-gradient(135deg, rgba(15,42,74,.10) 0%, rgba(56,189,248,.14) 100%);
  --accent-solid:#1D4ED8;

  /* Sombras frías difusas */
  --shadow-xs:0 1px 3px rgba(30,40,60,.04);
  --shadow-sm:0 2px 8px rgba(30,40,60,.06);
  --shadow-md:0 4px 16px rgba(30,40,60,.08);
  --shadow-lg:0 8px 32px rgba(30,40,60,.10);
  --shadow-xl:0 16px 48px rgba(30,40,60,.13);
  --shadow-inset:inset 0 1px 0 rgba(255,255,255,.7);
  --glass:rgba(255,255,255,.12);
  --glass-border:rgba(255,255,255,.55);
  --blur:blur(5px);
}
.cc-root{
  font-family:'Montserrat',-apple-system,sans-serif;
  font-weight:300;
  font-size:14px;
  color:var(--ink); background:transparent;
  min-height:100vh; width:100%;
  position:relative;
  -webkit-font-smoothing:antialiased;
  position:relative;
}
.cc-bg-wave{display:none;}


.cc-num{font-variant-numeric:tabular-nums;font-feature-settings:"tnum";}
.cc-emoji{}
.cc-serif{font-family:'Fraunces',serif;letter-spacing:-.018em;font-feature-settings:"ss01";}

.cc-wrap{max-width:760px;margin:0 auto;padding:4px 16px 130px;}

/* ============== TOPBAR ============== */
.cc-top{position:sticky;top:0;z-index:30;
  background:transparent;
  padding:calc(14px + env(safe-area-inset-top)) 20px 8px;
  transition:.2s ease;}
.cc-top.scrolled{padding-top:calc(9px + env(safe-area-inset-top));padding-bottom:6px;
  background:rgba(255,255,255,.1);
  backdrop-filter:blur(3px);
  -webkit-backdrop-filter:blur(3px);
  border-bottom:none;}
.cc-top-inner{max-width:760px;margin:0 auto;}

/* logo centrado estilo Canva */
.cc-zafi-wordmark{font-family:'Fraunces',serif;font-weight:400;font-size:16px;
  letter-spacing:-.03em;text-align:center;color:var(--ink);
  font-feature-settings:"ss01"; margin-bottom:12px; transition:.2s;}
.cc-top.scrolled .cc-zafi-wordmark{margin-bottom:6px;font-size:14px;}

/* fila de perfil: avatar + nombre + plan */
.cc-profile-row{display:flex;align-items:center;gap:10px;margin-bottom:12px;transition:.2s;}
.cc-top.scrolled .cc-profile-row{margin-bottom:6px;}
.cc-avatar{width:38px;height:38px;border-radius:50%;flex-shrink:0;
  background:var(--ink);color:#fff;
  display:flex;align-items:center;justify-content:center;
  font-family:'Fraunces',serif;font-weight:600;font-size:15px;
  box-shadow:var(--shadow-sm);transition:.2s;}
.cc-top.scrolled .cc-avatar{width:30px;height:30px;font-size:13px;}
.cc-profile-name{font-family:'Fraunces',serif;font-weight:600;font-size:18px;
  letter-spacing:-.03em;line-height:1.2;color:var(--ink);transition:.2s;}
.cc-top.scrolled .cc-profile-name{font-size:15px;}
.cc-profile-plan{font-size:11px;color:var(--ink-soft);font-weight:500;letter-spacing:-.005em;}
.cc-top.scrolled .cc-profile-plan{display:none;}

.cc-masthead{display:flex;align-items:baseline;justify-content:space-between;gap:12px;
  margin-bottom:10px;transition:.2s;}
.cc-top.scrolled .cc-masthead{margin-bottom:6px;}
.cc-masthead-title{font-family:'Fraunces',serif;font-weight:600;font-size:26px;letter-spacing:-.045em;line-height:1;
  display:flex;align-items:center;gap:8px;font-feature-settings:"ss01";}
.cc-masthead-title::before{content:"";width:8px;height:8px;border-radius:50%;background:var(--green);}

/* Logo del onboarding */
.cc-logo{font-family:'Fraunces',serif;font-weight:400;font-size:26px;letter-spacing:-.05em;
  display:flex;align-items:center;gap:8px;color:var(--ink);font-feature-settings:"ss01";}
.cc-logo-dot{display:none;}
.cc-masthead-meta{font-size:10.5px;font-weight:500;color:var(--ink-faint);letter-spacing:.02em;
  font-variant-numeric:tabular-nums;}

.cc-balance-row{display:flex;align-items:flex-end;justify-content:space-between;gap:14px;padding:2px 0 14px;
  transition:.2s;}
.cc-top.scrolled .cc-balance-row{padding:0 0 8px;}
.cc-top.scrolled .cc-balance-value{font-size:20px;}
.cc-balance-display{display:flex;flex-direction:column;gap:2px;min-width:0;}
.cc-balance-label{font-size:11px;font-weight:600;color:var(--ink-soft);letter-spacing:-.005em;}
.cc-balance-value{font-family:'Fraunces',serif;font-weight:500;font-size:32px;line-height:1;letter-spacing:-.035em;
  font-variant-numeric:tabular-nums;transition:font-size .2s;}

/* chip de rango */
.cc-range-chip{display:inline-flex;align-items:center;gap:7px;padding:8px 14px;
  background:var(--glass);border:1px solid var(--glass-border);border-radius:14px;
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  font-family:inherit;font-size:12.5px;font-weight:600;color:var(--ink);cursor:pointer;
  transition:all .15s ease;}
.cc-range-chip:hover{background:rgba(255,255,255,.7);}
.cc-range-chip:active{transform:scale(.98);}
.cc-range-chip .cc-range-arrow{color:var(--ink-faint);font-size:9px;}

/* tabs — segmented control glass (legado, ya no se usa en header) */
.cc-tabs{display:flex;gap:2px;background:rgba(0,0,0,.04);
  border:1px solid var(--glass-border);border-radius:16px;padding:3px;
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  margin-top:4px;}
.cc-tab{flex:1;font-family:inherit;font-size:12.5px;font-weight:600;
  border:none;background:transparent;color:var(--ink-faint);
  padding:8px 6px;border-radius:13px;cursor:pointer;
  transition:all .15s ease;position:relative;}
.cc-tab:hover{color:var(--ink);}
.cc-tab.on{background:rgba(255,255,255,.1);color:var(--ink);
  box-shadow:var(--shadow-sm);}

/* ============== TARJETAS ============== */
.cc-card{background:var(--glass);backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  border:1px solid var(--glass-border);
  border-radius:20px;padding:0;box-shadow:var(--shadow-sm);
  transition:.2s ease;}
.cc-card-boxed{background:var(--glass);backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  border:1px solid var(--glass-border);
  border-radius:20px;padding:16px;box-shadow:var(--shadow-sm);}
.cc-card-section{background:transparent;border:none;border-radius:0;padding-top:8px;padding-bottom:4px;
  box-shadow:none;}
.cc-fade{animation:ccUp .4s cubic-bezier(.16,1,.3,1) both;}
@keyframes ccUp{from{opacity:0;}to{opacity:1;}}

/* ============== BOTONES ============== */
.cc-btn{font-family:inherit;font-size:13.5px;font-weight:600;border-radius:14px;border:1px solid var(--glass-border);
  background:var(--glass);backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  color:var(--ink);padding:10px 17px;cursor:pointer;
  transition:all .15s ease;box-shadow:none;}
.cc-btn:hover{background:rgba(255,255,255,.7);}
.cc-btn:active{transform:scale(.98);}
.cc-btn-primary{background:var(--ink);color:#fff;border-color:transparent;backdrop-filter:none;-webkit-backdrop-filter:none;}
.cc-btn-primary:hover{background:#222;}
.cc-btn-green{background:var(--ink);color:#fff;
  border-color:transparent;backdrop-filter:none;-webkit-backdrop-filter:none;}
.cc-btn-green:hover{background:#222;}
.cc-btn-green:active{transform:scale(.98);}
.cc-btn:disabled{opacity:.4;cursor:not-allowed;transform:none;}
.cc-btn:disabled:hover{background:var(--paper);border-color:var(--line);}

/* ============== INPUTS ============== */
.cc-input,.cc-select{font-family:inherit;font-size:14.5px;width:100%;padding:12px 16px;border-radius:16px;
  border:1px solid var(--glass-border);background:var(--glass);color:var(--ink);outline:none;
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  transition:border-color .15s ease;}
.cc-input:focus,.cc-select:focus{border-color:rgba(0,0,0,.15);background:rgba(255,255,255,.7);}
.cc-label{font-size:11.5px;font-weight:600;color:var(--ink-soft);margin-bottom:6px;display:block;letter-spacing:-.005em;}

/* ============== CHIPS — píldoras suaves ============== */
.cc-chip{font-family:inherit;font-size:12.5px;font-weight:600;
  border:1px solid var(--line);background:var(--paper);color:var(--ink);
  padding:7px 13px;border-radius:999px;cursor:pointer;
  transition:all .2s cubic-bezier(.2,.7,.2,1);
  box-shadow:var(--shadow-xs);}
.cc-chip:hover{border-color:var(--gold);transform:translateY(-1px);
  box-shadow:var(--shadow-sm);}

/* ============== BOTTOM NAV ============== */
.cc-bottomnav{position:fixed;left:0;right:0;bottom:0;z-index:50;
  display:flex;justify-content:center;
  padding:0 16px calc(14px + env(safe-area-inset-bottom));
  pointer-events:none;}
.cc-bottomnav-inner{pointer-events:auto;
  width:100%;max-width:420px;height:62px;
  display:flex;align-items:center;justify-content:space-between;
  background:rgba(255,255,255,.08);border:1px solid var(--glass-border);
  border-radius:31px;padding:0 10px;
  backdrop-filter:blur(3px);-webkit-backdrop-filter:blur(3px);
  box-shadow:var(--shadow-lg);
  position:relative;}
.cc-nav-item{flex:1;display:flex;flex-direction:column;align-items:center;gap:3px;
  border:none;background:transparent;cursor:pointer;
  padding:8px 4px 6px;border-radius:18px;
  color:var(--ink-faint);font-family:inherit;
  transition:color .15s ease, background .15s ease;}
.cc-nav-item.on{color:var(--ink);}
.cc-nav-item:hover{color:var(--ink);}
.cc-nav-icon{display:flex;align-items:center;justify-content:center;line-height:0;}
.cc-nav-icon svg{width:21px;height:21px;display:block;}
.cc-nav-label{font-size:10px;font-weight:600;letter-spacing:.01em;}

/* Orb central de IA — protruye arriba y abajo de la barra, centrado */
.cc-orb-slot{flex:0 0 78px;display:flex;align-items:center;justify-content:center;position:relative;height:100%;}
.cc-orb-btn{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);
  width:78px;height:78px;border-radius:50%;border:none;cursor:pointer;
  display:flex;align-items:center;justify-content:center;
  box-shadow:var(--shadow-lg);
  transition:transform .2s cubic-bezier(.34,1.56,.64,1), box-shadow .2s ease, opacity .2s ease;}
.cc-orb-btn:active{transform:translate(-50%,-50%) scale(.93);}
.cc-orb{width:78px;height:78px;border-radius:50%;position:relative;
  background:
    radial-gradient(circle at 35% 75%, rgba(99,102,241,.95) 0%, rgba(99,102,241,0) 45%),
    radial-gradient(circle at 70% 55%, rgba(96,165,250,.9) 0%, rgba(96,165,250,0) 50%),
    radial-gradient(circle at 55% 25%, rgba(94,234,212,.55) 0%, rgba(94,234,212,0) 40%),
    radial-gradient(circle at 25% 30%, rgba(167,139,250,.85) 0%, rgba(167,139,250,0) 55%),
    radial-gradient(circle at 50% 50%, #7C8BF5 0%, #5B6EE8 100%);
  filter:saturate(1.25);
  border:2px solid rgba(255,255,255,.55);
  box-shadow:inset -6px -8px 16px rgba(40,30,90,.35), inset 6px 6px 14px rgba(255,255,255,.35);
  animation:ccOrbBreathe 4s ease-in-out infinite, ccOrbDrift 8s ease-in-out infinite;}
.cc-orb::before{content:"";position:absolute;inset:0;border-radius:50%;
  background:radial-gradient(circle at 30% 28%, rgba(255,255,255,.85) 0%, rgba(255,255,255,0) 28%);
  animation:ccOrbShine 4s ease-in-out infinite;}
.cc-orb::after{content:"";position:absolute;inset:-10px;border-radius:50%;z-index:-1;
  background:radial-gradient(circle, rgba(124,139,245,.55) 0%, rgba(167,139,250,.25) 45%, rgba(124,139,245,0) 70%);
  animation:ccOrbGlow 4s ease-in-out infinite;}
@keyframes ccOrbBreathe{0%,100%{transform:scale(1);}50%{transform:scale(1.07);}}
@keyframes ccOrbDrift{
  0%,100%{background-position:35% 75%, 70% 55%, 55% 25%, 25% 30%, 50% 50%;filter:saturate(1.25) hue-rotate(0deg);}
  33%{background-position:55% 65%, 45% 70%, 70% 35%, 35% 45%, 50% 50%;filter:saturate(1.4) hue-rotate(-12deg);}
  66%{background-position:25% 60%, 75% 45%, 40% 30%, 50% 25%, 50% 50%;filter:saturate(1.35) hue-rotate(14deg);}}
@keyframes ccOrbShine{0%,100%{opacity:.85;transform:scale(1);}50%{opacity:1;transform:scale(1.04) translate(2px,2px);}}
@keyframes ccOrbGlow{0%,100%{opacity:.6;transform:scale(1);}50%{opacity:1;transform:scale(1.12);}}

/* ============== FAB superior (+) ============== */
.cc-fab-top{position:fixed;top:18px;right:18px;z-index:45;
  width:46px;height:46px;border-radius:50%;
  background:var(--glass);color:var(--ink);border:1px solid var(--glass-border);
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  font-size:24px;font-weight:300;line-height:1;cursor:pointer;
  display:flex;align-items:center;justify-content:center;
  box-shadow:var(--shadow-md);
  transition:all .15s ease;}
.cc-fab-top:hover{background:rgba(255,255,255,.75);box-shadow:var(--shadow-lg);}
.cc-fab-top:active{transform:scale(.93);}
.cc-fab-top.open{transform:rotate(45deg);}
.cc-fab-menu{position:fixed;top:70px;right:18px;z-index:45;display:flex;flex-direction:column;gap:8px;
  align-items:flex-end;animation:ccUp .15s cubic-bezier(.16,1,.3,1);}
.cc-fab-mini{font-family:inherit;font-size:13px;font-weight:600;padding:10px 16px;border-radius:10px;
  background:var(--paper);color:var(--ink);border:1px solid var(--line);cursor:pointer;
  box-shadow:var(--shadow-md);
  display:flex;align-items:center;gap:8px;
  transition:all .15s ease;}
.cc-fab-mini:hover{background:var(--surface-2);}

/* ============== TARJETAS DE CUENTAS ============== */
.cc-acc-card{cursor:pointer;
  border:1px solid var(--glass-border);background:var(--glass);
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  border-radius:18px;padding:13px 14px;min-width:148px;text-align:left;
  position:relative;overflow:hidden;
  transition:all .2s ease;}
.cc-acc-card:hover{background:rgba(255,255,255,.7);}
.cc-acc-card.on{border-color:rgba(29,78,216,.3);
  background:var(--accent-grad-soft);
  box-shadow:0 0 0 1px rgba(29,78,216,.2);}
.cc-acc-card.on .cc-acc-label{color:var(--accent-solid);}
.cc-acc-card.on .cc-acc-icon{background:rgba(29,78,216,.12);}
.cc-acc-icon{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;
  border-radius:7px;background:var(--surface-2);font-size:13px;margin-bottom:7px;
  transition:.2s;}
.cc-acc-label{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.07em;color:var(--ink-faint);}
.cc-acc-name{font-weight:600;font-size:14px;margin:2px 0 7px;letter-spacing:-.012em;color:var(--ink);}
.cc-acc-bal{font-family:'Fraunces',serif;font-weight:500;font-size:19px;letter-spacing:-.03em;line-height:1.05;}
.cc-grad-text{background:var(--accent-grad);-webkit-background-clip:text;background-clip:text;
  color:transparent;-webkit-text-fill-color:transparent;}
.cc-acc-sub{font-size:10.5px;color:var(--ink-soft);margin-top:4px;font-variant-numeric:tabular-nums;font-weight:500;}
.cc-scroll-x{display:flex;gap:8px;overflow-x:auto;padding:4px 2px 10px;scrollbar-width:none;}
.cc-scroll-x::-webkit-scrollbar{display:none;}

/* configurar */
.cc-gear{background:var(--glass);border:1px solid var(--glass-border);border-radius:14px;
  backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  padding:8px 14px;cursor:pointer;font-size:12.5px;font-weight:600;color:var(--ink-soft);
  display:inline-flex;align-items:center;gap:6px;
  transition:.15s ease;}
.cc-gear:hover{background:rgba(255,255,255,.7);color:var(--ink);}
.cc-gear svg, .cc-range-chip svg{width:14px;height:14px;flex-shrink:0;}

/* fila draggable */
.cc-sortable{padding:11px 13px;border:1px solid var(--line);border-radius:12px;
  background:var(--paper);display:flex;align-items:center;gap:10px;cursor:grab;
  user-select:none;transition:.18s;box-shadow:var(--shadow-xs);}
.cc-sortable:hover{border-color:var(--gold);box-shadow:var(--shadow-sm);}
.cc-sortable.disabled{opacity:.5;background:var(--surface-2);}
.cc-sortable .cc-grip-h{color:var(--ink-soft);font-size:18px;line-height:1;}

/* ============== SEPARADOR DE DÍA ============== */
.cc-day-sep{display:flex;align-items:center;justify-content:space-between;gap:12px;
  padding:13px 0 8px;border-bottom:1px solid var(--line-soft);margin-bottom:2px;
  position:sticky;top:0;background:transparent;
  z-index:5;}
.cc-day-sep:first-child{padding-top:6px;}
.cc-day-num{font-family:'Fraunces',serif;font-weight:500;font-size:19px;line-height:1;letter-spacing:-.03em;
  color:var(--ink);min-width:24px;}
.cc-day-name{flex:1;font-size:10.5px;font-weight:600;color:var(--ink-faint);letter-spacing:.04em;
  text-transform:uppercase;}
.cc-day-totals{display:flex;gap:9px;align-items:center;font-size:11px;
  font-variant-numeric:tabular-nums;font-weight:500;opacity:.7;}
.cc-day-totals .pos{color:var(--green);}
.cc-day-totals .neg{color:var(--coral);}

/* ============== MODAL ============== */
.cc-overlay{position:fixed;inset:0;background:rgba(0,0,0,.25);
  backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);
  z-index:10000;display:flex;align-items:flex-end;justify-content:center;
  animation:ccFadeIn .15s ease;}
@keyframes ccFadeIn{from{opacity:0;}to{opacity:1;}}
.cc-sheet{background:rgba(255,255,255,.7);backdrop-filter:blur(5px);-webkit-backdrop-filter:blur(5px);
  border-radius:24px 24px 0 0;width:100%;max-width:760px;
  max-height:92vh;overflow-y:auto;padding:10px 20px 28px;
  animation:ccSheet .3s cubic-bezier(.16,1,.3,1);
  border-top:1px solid rgba(255,255,255,.6);
  box-shadow:0 -4px 24px rgba(0,0,0,.08);}
@keyframes ccSheet{from{transform:translateY(100%);}to{transform:none;}}
.cc-grip{width:36px;height:4px;background:rgba(0,0,0,.12);border-radius:99px;margin:8px auto 16px;}
.cc-sheet-top{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;}
.cc-sheet-top h2{font-family:'Fraunces',serif;font-size:21px;font-weight:600;color:var(--ink);}
.cc-sheet-close{width:32px;height:32px;border-radius:50%;border:none;background:var(--surface);
  display:flex;align-items:center;justify-content:center;cursor:pointer;color:var(--ink-soft);
  font-size:18px;line-height:1;transition:.15s ease;flex-shrink:0;}
.cc-sheet-close:hover{background:var(--surface-2);color:var(--ink);}

/* monto grande centrado — estilo "Nueva transacción" */
.cc-amount-display{display:flex;align-items:baseline;justify-content:center;gap:6px;
  padding:18px 0;margin-bottom:6px;width:100%;}
.cc-amount-display .cc-amount-currency{font-family:'Fraunces',serif;font-size:28px;font-weight:500;color:var(--ink-soft);flex-shrink:0;}
.cc-amount-display .cc-amount-mxn{font-family:'Montserrat',sans-serif;font-size:16px;font-weight:300;color:var(--ink-faint);flex-shrink:0;align-self:center;}
.cc-amount-display input{font-family:'Fraunces',serif;font-size:42px;font-weight:600;letter-spacing:-.02em;
  text-align:left;color:var(--ink);background:transparent;border:none;outline:none;
  width:auto;max-width:240px;min-width:30px;font-feature-settings:"tnum";
  -moz-appearance:textfield;appearance:textfield;}
.cc-amount-display input::-webkit-outer-spin-button,
.cc-amount-display input::-webkit-inner-spin-button{-webkit-appearance:none;-moz-appearance:none;appearance:none;margin:0;display:none;}
.cc-amount-display input[type=number]{-moz-appearance:textfield;}
.cc-amount-display input::placeholder{color:var(--ink-soft);opacity:.55;}

/* ============== CHAT ============== */
.cc-bubble{padding:12px 15px;border-radius:18px;font-size:14.5px;line-height:1.5;max-width:84%;
  letter-spacing:-.005em;}
.cc-bubble.bot{background:rgba(255,255,255,.55);backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  color:var(--ink);border:1px solid var(--glass-border);
  border-bottom-left-radius:4px;}
.cc-bubble.me{background:rgba(255,255,255,.45);backdrop-filter:var(--blur);-webkit-backdrop-filter:var(--blur);
  color:var(--ink);border:1px solid var(--glass-border);
  border-bottom-right-radius:6px;align-self:flex-end;
  box-shadow:var(--shadow-sm);}
.cc-mic.rec{background:var(--coral)!important;color:#fff!important;border-color:var(--coral)!important;
  animation:ccMicPulse 1.1s ease-in-out infinite;}
@keyframes ccMicPulse{0%,100%{box-shadow:0 0 0 0 rgba(181,69,58,.5);}50%{box-shadow:0 0 0 8px rgba(181,69,58,0);}}

/* ===== Splash ===== */
.cc-splash{position:fixed;inset:0;z-index:99999;display:flex;flex-direction:column;
  align-items:center;justify-content:center;gap:10px;
  background:
    radial-gradient(circle at 50% 38%, #EAEEF4 0%, #DCE1E8 55%, #D2D8E2 100%);
  animation:ccSplashOut .5s ease 1.85s forwards;}
@keyframes ccSplashOut{to{opacity:0;visibility:hidden;}}
.cc-splash-orb{position:absolute;top:calc(50% - 70px);left:50%;
  width:120px;height:120px;border-radius:50%;transform:translate(-50%,-50%) scale(.4);
  background:
    radial-gradient(circle at 35% 75%, rgba(99,102,241,.9) 0%, rgba(99,102,241,0) 45%),
    radial-gradient(circle at 70% 55%, rgba(96,165,250,.85) 0%, rgba(96,165,250,0) 50%),
    radial-gradient(circle at 55% 25%, rgba(94,234,212,.5) 0%, rgba(94,234,212,0) 40%),
    radial-gradient(circle at 25% 30%, rgba(167,139,250,.8) 0%, rgba(167,139,250,0) 55%),
    radial-gradient(circle at 50% 50%, #7C8BF5 0%, #5B6EE8 100%);
  filter:blur(8px);opacity:0;
  animation:ccSplashOrb 1.7s cubic-bezier(.2,.7,.2,1) forwards, ccOrbDrift 8s ease-in-out infinite;}
@keyframes ccSplashOrb{
  0%{opacity:0;transform:translate(-50%,-50%) scale(.3);}
  40%{opacity:.85;transform:translate(-50%,-50%) scale(1.05);}
  100%{opacity:.6;transform:translate(-50%,-50%) scale(1);}}
.cc-splash-word{position:relative;font-family:'Fraunces',serif;font-weight:400;
  font-size:54px;letter-spacing:.1em;color:var(--ink);
  opacity:0;filter:blur(16px);
  animation:ccSplashBlur 1s ease .35s forwards;}
@keyframes ccSplashBlur{to{opacity:1;filter:blur(0);letter-spacing:-.05em;}}
.cc-splash-tag{position:relative;font-family:'Montserrat',sans-serif;font-weight:300;
  font-size:13px;letter-spacing:.22em;text-transform:uppercase;color:var(--ink-soft);
  opacity:0;animation:ccSplashTag .7s ease .95s forwards;}
@keyframes ccSplashTag{to{opacity:1;}}

.cc-toast{position:fixed;left:50%;transform:translateX(-50%);bottom:96px;z-index:60;
  background:linear-gradient(160deg,#2c2820,var(--ink));color:var(--paper);
  padding:12px 20px;border-radius:14px;font-size:13.5px;font-weight:600;letter-spacing:-.005em;
  box-shadow:0 12px 32px rgba(31,27,20,.32),var(--shadow-inset);
  animation:ccUp .3s cubic-bezier(.2,.7,.2,1);}

.cc-dots span{display:inline-block;width:6px;height:6px;border-radius:50%;background:var(--ink-soft);margin:0 2px;
  animation:ccDot 1s infinite;}
.cc-dots span:nth-child(2){animation-delay:.15s;}
.cc-dots span:nth-child(3){animation-delay:.3s;}
@keyframes ccDot{0%,60%,100%{opacity:.25;transform:translateY(0);}30%{opacity:1;transform:translateY(-3px);}}

/* ============== ENCABEZADOS DE SECCIÓN ============== */
.cc-eyebrow{font-size:11px;font-weight:600;color:var(--ink-soft);letter-spacing:.02em;
  display:flex;align-items:center;gap:8px;margin-bottom:5px;}
.cc-section-title{font-family:'Fraunces',serif;font-weight:500;font-size:22px;
  letter-spacing:-.03em;line-height:1.15;margin-bottom:14px;}

/* selection */
::selection{background:var(--gold-glow);color:var(--ink);}
`;

/* --------------------------- datos por defecto --------------------------- */
const DEFAULT_CATS = [
  { name: "Sueldo", emoji: "💼", type: "income" },
  { name: "Negocio", emoji: "🏢", type: "income" },
  { name: "Freelance", emoji: "💻", type: "income" },
  { name: "Otros ingresos", emoji: "💰", type: "income" },
  { name: "Súper / Despensa", emoji: "🛒", type: "expense" },
  { name: "Restaurantes", emoji: "🍔", type: "expense" },
  { name: "Transporte / Gasolina", emoji: "⛽", type: "expense" },
  { name: "Casa / Renta", emoji: "🏠", type: "expense" },
  { name: "Servicios", emoji: "💡", type: "expense" },
  { name: "Salud", emoji: "🏥", type: "expense" },
  { name: "Entretenimiento", emoji: "🎬", type: "expense" },
  { name: "Suscripciones", emoji: "📱", type: "expense" },
  { name: "Otros gastos", emoji: "📦", type: "expense" },
];

const SEED_KW = {
  "sueldo": ["sueldo", "salario", "nomina", "quincena"],
  "negocio": ["venta", "cliente", "factura"],
  "freelance": ["freelance", "proyecto", "honorarios"],
  "super / despensa": ["super", "despensa", "walmart", "soriana", "calimax", "mercado", "verdura", "fruta", "leche"],
  "restaurantes": ["restaurante", "taco", "tacos", "pizza", "comida", "cena", "almuerzo", "cafe", "starbucks", "sushi", "hamburguesa"],
  "transporte / gasolina": ["gasolina", "uber", "didi", "taxi", "pasaje", "caseta", "estacionamiento", "pemex"],
  "casa / renta": ["renta", "hipoteca", "mantenimiento"],
  "servicios": ["luz", "cfe", "agua", "internet", "telmex", "izzi", "totalplay", "telefono"],
  "salud": ["doctor", "medicina", "farmacia", "consulta", "dentista", "hospital"],
  "entretenimiento": ["cine", "concierto", "boletos", "antro", "juego"],
  "suscripciones": ["netflix", "spotify", "disney", "suscripcion", "prime", "youtube"],
};

/* ------------------------------ utilidades ------------------------------- */
const uid = () => Math.random().toString(36).slice(2, 10);
const today = () => new Date().toISOString().slice(0, 10);

/* ===== Sistema de idiomas (i18n) ===== */
let _lang = "es";
const setAppLang = (l) => { _lang = l; };
const STRINGS = {
  // Tabs
  home: { es: "Inicio", en: "Home" },
  history: { es: "Historial", en: "History" },
  categories: { es: "Categorías", en: "Categories" },
  statistics: { es: "Estadísticas", en: "Statistics" },
  // Header
  weeklyPlan: { es: "Plan semanal", en: "Weekly plan" },
  // Home sections
  yourAccounts: { es: "Tus cuentas", en: "Your accounts" },
  addAccount: { es: "Agregar cuenta", en: "Add account" },
  customize: { es: "Personalizar", en: "Customize" },
  incomeThisMonth: { es: "Ingresos · Este mes", en: "Income · This month" },
  expenseThisMonth: { es: "Gastos · Este mes", en: "Expenses · This month" },
  topExpenses: { es: "Gastos más grandes", en: "Biggest expenses" },
  recentMovements: { es: "Últimos movimientos", en: "Recent movements" },
  expByCategory: { es: "Gastos por categoría", en: "Expenses by category" },
  balance: { es: "Saldo", en: "Balance" },
  // History
  all: { es: "Todos", en: "All" },
  income: { es: "Ingresos", en: "Income" },
  expenses: { es: "Gastos", en: "Expenses" },
  select: { es: "Seleccionar", en: "Select" },
  selected: { es: "seleccionado", en: "selected" },
  selectAll: { es: "Todos", en: "All" },
  none: { es: "Ninguno", en: "None" },
  deleteBtn: { es: "Eliminar", en: "Delete" },
  sortBy: { es: "Ordenar", en: "Sort by" },
  mov: { es: "mov.", en: "mov." },
  touchToEdit: { es: "Toca un movimiento para editarlo · ✕ para eliminarlo", en: "Tap a movement to edit it · ✕ to delete it" },
  noMovements: { es: "No hay movimientos en este periodo.", en: "No movements in this period." },
  // AddModal
  newTransaction: { es: "Nueva transacción", en: "New transaction" },
  editTransaction: { es: "Editar transacción", en: "Edit transaction" },
  expense: { es: "Gasto", en: "Expense" },
  incomeType: { es: "Ingreso", en: "Income" },
  concept: { es: "Concepto", en: "Concept" },
  date: { es: "Fecha", en: "Date" },
  category: { es: "Categoría", en: "Category" },
  account: { es: "Cuenta", en: "Account" },
  save: { es: "Guardar", en: "Save" },
  saveChanges: { es: "Guardar cambios", en: "Save changes" },
  // TopFab menu
  addMovement: { es: "Agregar movimiento", en: "Add movement" },
  manualCapture: { es: "Capturar manual", en: "Manual capture" },
  manualCaptureDesc: { es: "Escribe el movimiento tú mismo", en: "Enter the movement yourself" },
  recurringMovement: { es: "Movimiento recurrente", en: "Recurring movement" },
  recurringDesc: { es: "Se repite automáticamente", en: "Repeats automatically" },
  fromScreenshot: { es: "Desde screenshot", en: "From screenshot" },
  fromScreenshotDesc: { es: "Sube una captura y la IA lo lee", en: "Upload a screenshot and AI reads it" },
  fromExcel: { es: "Desde Excel", en: "From Excel" },
  fromExcelDesc: { es: "Importa una hoja de cálculo", en: "Import a spreadsheet" },
  // Settings
  settings: { es: "Ajustes", en: "Settings" },
  personalInfo: { es: "Información personal", en: "Personal information" },
  name: { es: "Nombre", en: "Name" },
  phone: { es: "Teléfono", en: "Phone" },
  email: { es: "Correo electrónico", en: "Email" },
  language: { es: "Idioma", en: "Language" },
  currency: { es: "Moneda", en: "Currency" },
  notifications: { es: "Notificaciones", en: "Notifications" },
  expenseReminders: { es: "Recordatorios de gastos", en: "Expense reminders" },
  comingSoon: { es: "Próximamente", en: "Coming soon" },
  dataPrivacy: { es: "Datos y privacidad", en: "Data & privacy" },
  exportData: { es: "Exportar mis datos", en: "Export my data" },
  resetApp: { es: "Reiniciar app (borrar todo)", en: "Reset app (delete all)" },
  resetConfirm: { es: "¿Estás seguro? Esto borrará todas tus categorías, cuentas y movimientos.", en: "Are you sure? This will delete all your categories, accounts and movements." },
  cancel: { es: "Cancelar", en: "Cancel" },
  yesDeleteAll: { es: "Sí, borrar todo", en: "Yes, delete all" },
  signOut: { es: "Cerrar sesión", en: "Sign out" },
  signingOut: { es: "Cerrando sesión…", en: "Signing out…" },
  infoUpdated: { es: "Información actualizada", en: "Information updated" },
  chooseAvatar: { es: "Elige tu avatar", en: "Choose your avatar" },
  avatarUpdated: { es: "Avatar actualizado", en: "Avatar updated" },
  avatarRemoved: { es: "Avatar eliminado", en: "Avatar removed" },
  choosePhoto: { es: "Elegir foto", en: "Choose photo" },
  takePhoto: { es: "Tomar foto", en: "Take photo" },
  orUploadPhoto: { es: "O sube tu propia foto", en: "Or upload your own photo" },
  // Recurring
  recurringMovements: { es: "Movimientos recurrentes", en: "Recurring movements" },
  daily: { es: "Diario", en: "Daily" },
  weekly: { es: "Semanal", en: "Weekly" },
  biweekly: { es: "Quincenal", en: "Biweekly" },
  monthly: { es: "Mensual", en: "Monthly" },
  yearly: { es: "Anual", en: "Yearly" },
  manage: { es: "Gestionar", en: "Manage" },
  newRecurring: { es: "Nuevo recurrente", en: "New recurring" },
  pause: { es: "Pausar", en: "Pause" },
  activate: { es: "Activar", en: "Activate" },
  edit: { es: "Editar", en: "Edit" },
  // Stats
  summary: { es: "Resumen", en: "Summary" },
  netFlow: { es: "Flujo neto", en: "Net flow" },
  tapForDetail: { es: "Tocar para detalle ▸", en: "Tap for detail ▸" },
  avgDailyExpense: { es: "Gasto promedio diario", en: "Avg. daily expense" },
  movementsInPeriod: { es: "Movimientos en el periodo", en: "Movements in period" },
  incByCategory: { es: "Ingresos por categoría", en: "Income by category" },
  topSpent: { es: "En lo que más gastaste", en: "Top spending" },
  ofYourExpenses: { es: "de tus gastos", en: "of your expenses" },
  customizeStats: { es: "Personalizar", en: "Customize" },
  customizeStatsTitle: { es: "Personalizar estadísticas", en: "Customize statistics" },
  reorderHint: { es: "Reordena con las flechas y muestra u oculta cada sección.", en: "Reorder with arrows and show or hide each section." },
  bars: { es: "Barras", en: "Bars" },
  pie: { es: "Pastel", en: "Pie" },
  donut: { es: "Dona", en: "Donut" },
  slideToSee: { es: "Desliza sobre la gráfica para ver cualquier día", en: "Slide over the chart to see any day" },
  // Categories tab
  eachAccountHasOwn: { es: "Cada cuenta tiene sus propias categorías.", en: "Each account has its own categories." },
  totalsAreFrom: { es: "Los totales son de", en: "Totals are from" },
  noRecurring: { es: "No tienes movimientos recurrentes. Crea uno para que se registre automáticamente.", en: "No recurring movements. Create one to have it registered automatically." },
  activeCount: { es: "activo", en: "active" },
  autoGenerated: { es: "se generan solos en su fecha", en: "auto-generated on their date" },
  addCategoryTo: { es: "Agregar categoría a", en: "Add category to" },
  noCats: { es: "Sin categorías.", en: "No categories." },
  inPeriod: { es: "en el periodo", en: "in period" },
  // Assistant
  assistant: { es: "Asistente", en: "Assistant" },
  close: { es: "Cerrar", en: "Close" },
  send: { es: "Enviar", en: "Send" },
  releaseToSend: { es: "Suelta para enviar", en: "Release to send" },
  listening: { es: "Escuchando…", en: "Listening…" },
  // General
  loading: { es: "Cargando…", en: "Loading…" },
  user: { es: "Usuario", en: "User" },
  from: { es: "Desde", en: "From" },
  to: { es: "Hasta", en: "To" },
  apply: { es: "Aplicar", en: "Apply" },
  male: { es: "Hombre", en: "Male" },
  female: { es: "Mujer", en: "Female" },
  other: { es: "Otro", en: "Other" },
  paused: { es: "pausado", en: "paused" },
};
const t = (key) => (STRINGS[key] || {})[_lang] || (STRINGS[key] || {}).es || key;

/* ===== Motor de movimientos recurrentes =====
 * Una regla recurrente vive en config.recurring[] con forma:
 * { id, type:"expense"|"income", amount, description, accountId, categoryId,
 *   freq:"daily"|"weekly"|"biweekly"|"monthly"|"yearly", startDate:"YYYY-MM-DD",
 *   lastRun:"YYYY-MM-DD"|null, active:true }
 */
const FREQ_LABELS_FN = () => ({
  daily: t("daily"), weekly: t("weekly"), biweekly: t("biweekly"),
  monthly: t("monthly"), yearly: t("yearly"),
});
// Una regla está activa salvo que esté pausada explícitamente (compat: asistente usa `paused`, modal usa `active`)
const isRecActive = (r) => r.active !== false && r.paused !== true;

/* ===== Dictado por voz (Web Speech API) ===== */
const SpeechRec = typeof window !== "undefined"
  ? (window.SpeechRecognition || window.webkitSpeechRecognition)
  : null;
const voiceSupported = !!SpeechRec;
const isoToDate = (iso) => { const [y, m, d] = iso.split("-").map(Number); return new Date(y, m - 1, d); };
const dateToIso = (dt) => {
  const y = dt.getFullYear(); const m = String(dt.getMonth() + 1).padStart(2, "0"); const d = String(dt.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
};
// avanza una fecha según la frecuencia
const advanceDate = (dt, freq) => {
  const n = new Date(dt);
  if (freq === "daily") n.setDate(n.getDate() + 1);
  else if (freq === "weekly") n.setDate(n.getDate() + 7);
  else if (freq === "biweekly") n.setDate(n.getDate() + 14);
  else if (freq === "monthly") n.setMonth(n.getMonth() + 1);
  else if (freq === "yearly") n.setFullYear(n.getFullYear() + 1);
  return n;
};
// genera todos los movimientos pendientes de una regla, desde startDate (o lastRun) hasta hoy
const runRecurringRules = (recurring) => {
  if (!recurring || !recurring.length) return { newTxs: [], updatedRecurring: recurring || [] };
  const todayIso = today();
  const todayD = isoToDate(todayIso);
  const newTxs = [];
  const updatedRecurring = recurring.map((r) => {
    if (!isRecActive(r)) return r;
    // primera fecha pendiente
    let cursor;
    if (r.lastRun) {
      cursor = advanceDate(isoToDate(r.lastRun), r.freq);
    } else {
      cursor = isoToDate(r.startDate);
    }
    let lastRun = r.lastRun;
    let guard = 0;
    while (cursor <= todayD && guard < 1000) {
      const iso = dateToIso(cursor);
      newTxs.push({
        id: uid(),
        type: r.type,
        amount: r.amount,
        description: r.description,
        categoryId: r.categoryId || null,
        accountId: r.accountId,
        date: iso,
        recurringId: r.id,
      });
      lastRun = iso;
      cursor = advanceDate(cursor, r.freq);
      guard++;
    }
    return { ...r, lastRun };
  });
  return { newTxs, updatedRecurring };
};

const fmtBare = (n) => {
  const v = n || 0;
  const hasDecimals = v % 1 !== 0;
  return new Intl.NumberFormat("es-MX", { style: "currency", currency: "MXN",
    minimumFractionDigits: hasDecimals ? 2 : 0, maximumFractionDigits: hasDecimals ? 2 : 0 }).format(v);
};
const fmt = (n) => `${fmtBare(n)} mxn`;
const fmtMxn = (n) => fmt(n);
const norm = (s) =>
  (s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9 ]/g, " ");
const monthLabel = (mk) => {
  const [y, m] = mk.split("-");
  const names = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  return `${names[+m - 1]} ${y}`;
};
const DAY_NAMES = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
const MONTH_NAMES_FULL = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
const dayLabel = (isoDate) => {
  // YYYY-MM-DD → "jueves 19 de mayo 2026" sin desfase de zona horaria
  const [y, m, d] = isoDate.split("-").map(Number);
  if (!y || !m || !d) return isoDate;
  const date = new Date(y, m - 1, d);
  const todayD = new Date(today());
  const yest = new Date(todayD); yest.setDate(yest.getDate() - 1);
  const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  if (sameDay(date, todayD)) return "Hoy";
  if (sameDay(date, yest)) return "Ayer";
  return `${DAY_NAMES[date.getDay()]} ${d} de ${MONTH_NAMES_FULL[m - 1]} ${y}`;
};

// Para el estilo editorial: número del día grande + descripción a un lado
const dayParts = (isoDate) => {
  const [y, m, d] = isoDate.split("-").map(Number);
  if (!y || !m || !d) return { num: "—", desc: isoDate, isSpecial: false };
  const date = new Date(y, m - 1, d);
  const todayD = new Date(today());
  const yest = new Date(todayD); yest.setDate(yest.getDate() - 1);
  const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  if (sameDay(date, todayD)) return { num: String(d), desc: "HOY", isSpecial: true };
  if (sameDay(date, yest)) return { num: String(d), desc: "AYER", isSpecial: true };
  const mon = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"][m-1];
  const day = ["DOMINGO","LUNES","MARTES","MIÉRCOLES","JUEVES","VIERNES","SÁBADO"][date.getDay()];
  return { num: String(d), desc: `${day} · ${mon} ${y}`, isSpecial: false };
};
const STOP = new Set(["para", "con", "los", "las", "del", "una", "uno", "este", "esta", "que", "por", "mas",
  "muy", "pago", "compra", "gasto", "abono", "the", "and"]);

function extractKW(desc) {
  return norm(desc).split(/\s+/).filter((w) => w.length >= 4 && !STOP.has(w));
}

/* detección de duplicados: ¿existe ya un movimiento equivalente?
   - misma cuenta, mismo tipo, mismo monto (a 2 decimales)
   - fecha dentro de ±3 días (los bancos a veces ponen fecha de operación vs aplicación)
   - descripción "parecida": comparten al menos una palabra clave significativa,
     o las descripciones normalizadas son iguales */
function findDuplicate(candidate, existing) {
  const candDate = new Date(candidate.date).getTime();
  const candKws = new Set(extractKW(candidate.description));
  const candNorm = norm(candidate.description).trim();
  const candAmt = Math.round(candidate.amount * 100);
  for (const t of existing) {
    if (t.accountId !== candidate.accountId) continue;
    if (t.type !== candidate.type) continue;
    if (Math.round(t.amount * 100) !== candAmt) continue;
    const tDate = new Date(t.date).getTime();
    const diffDays = Math.abs((tDate - candDate) / 86400000);
    if (diffDays > 3) continue;
    const tNorm = norm(t.description).trim();
    if (candNorm && tNorm && candNorm === tNorm) return t;
    const tKws = extractKW(t.description);
    if (tKws.some((kw) => candKws.has(kw))) return t;
    // si una descripción está vacía pero monto+fecha+cuenta+tipo coinciden, también es duplicado probable
    if (!candNorm || !tNorm) return t;
  }
  return null;
}

/* saldos: saldo inicial + movimientos */
function accountBalance(config, txs, accId) {
  const acc = config.accounts.find((a) => a.id === accId);
  const init = acc ? acc.initialBalance || 0 : 0;
  return init + txs.filter((t) => t.accountId === accId)
    .reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
}
function grandTotal(config, txs) {
  const init = config.accounts.reduce((s, a) => s + (a.initialBalance || 0), 0);
  return init + txs.reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
}

/* Para estadísticas: separa entre movimientos reales (afectan ingresos/gastos)
   y pass-through (dinero de paso). Los pass-through con el mismo groupId
   forman un grupo: sumamos ingresos del grupo, restamos gastos del grupo,
   la diferencia se cuenta como ingreso/gasto real (la "propina" o "lo que pusiste").
   Soporta cualquier combinación: 1-1, 1-n, n-1, n-m. */
function statTxs(txs) {
  const real = [];
  const synthetic = [];
  const used = new Set();
  // 1. agrupar pass-through por groupId
  const groups = new Map();
  for (const t of txs) {
    if (!t.passThrough) { real.push(t); continue; }
    if (!t.groupId) { used.add(t.id); continue; } // pass-through sin grupo → ignorar en stats
    if (!groups.has(t.groupId)) groups.set(t.groupId, []);
    groups.get(t.groupId).push(t);
  }
  // 2. para cada grupo: sumar ingresos vs gastos
  for (const [gid, members] of groups) {
    let totalIn = 0, totalOut = 0;
    members.forEach((m) => {
      used.add(m.id);
      if (m.type === "income") totalIn += m.amount;
      else totalOut += m.amount;
    });
    const diff = totalIn - totalOut;
    if (Math.abs(diff) < 0.01) continue;
    // fecha del sintético: la más reciente del grupo
    const latest = members.reduce((a, b) => (a.date >= b.date ? a : b));
    // descripción del sintético: tomar la de algún miembro
    const sampleDesc = members.find((m) => m.description)?.description || "reembolso";
    synthetic.push({
      id: "synth-" + gid,
      type: diff >= 0 ? "income" : "expense",
      amount: Math.abs(diff),
      description: `Diferencia neta · ${sampleDesc}`,
      categoryId: null,
      accountId: members[0].accountId,
      date: latest.date,
    });
  }
  return { real, synthetic, all: [...real, ...synthetic] };
}

/* ===================== RANGO TEMPORAL GLOBAL ============================ */
/* El rango global vive en config.dateRange con forma:
     { preset: "today"|"week"|"month"|"last-month"|"3m"|"6m"|"year"|"last-year"|"all"|"custom",
       from?: "YYYY-MM-DD", to?: "YYYY-MM-DD" }
   Si preset !== "custom", from/to se calculan dinámicamente con resolveRange().
*/
const DEFAULT_RANGE = { preset: "month" };

function resolveRange(range) {
  const r = range || DEFAULT_RANGE;
  const t = today();
  const td = new Date(t + "T12:00:00");
  const iso = (d) => d.toISOString().slice(0, 10);
  function firstOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1, 12); }
  function lastOfMonth(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 0, 12); }
  switch (r.preset) {
    case "today": return { from: t, to: t };
    case "week": {
      const dow = td.getDay();
      const monOff = dow === 0 ? -6 : 1 - dow;
      const mon = new Date(td); mon.setDate(mon.getDate() + monOff);
      const sun = new Date(mon); sun.setDate(sun.getDate() + 6);
      return { from: iso(mon), to: iso(sun) };
    }
    case "month": return { from: iso(firstOfMonth(td)), to: t };
    case "last-month": {
      const prev = new Date(td.getFullYear(), td.getMonth() - 1, 15, 12);
      return { from: iso(firstOfMonth(prev)), to: iso(lastOfMonth(prev)) };
    }
    case "3m": {
      const start = new Date(td.getFullYear(), td.getMonth() - 2, 1, 12);
      return { from: iso(start), to: t };
    }
    case "6m": {
      const start = new Date(td.getFullYear(), td.getMonth() - 5, 1, 12);
      return { from: iso(start), to: t };
    }
    case "year": return { from: `${td.getFullYear()}-01-01`, to: t };
    case "last-year": {
      const y = td.getFullYear() - 1;
      return { from: `${y}-01-01`, to: `${y}-12-31` };
    }
    case "all": return { from: "1970-01-01", to: t };
    case "custom":
      return { from: r.from || t, to: r.to || t };
    default: return { from: iso(firstOfMonth(td)), to: t };
  }
}

function rangeLabel(range) {
  const r = range || DEFAULT_RANGE;
  if (r.preset === "today") return "Hoy";
  if (r.preset === "week") return "Esta semana";
  if (r.preset === "month") return "Este mes";
  if (r.preset === "last-month") return "Mes pasado";
  if (r.preset === "3m") return "Últimos 3 meses";
  if (r.preset === "6m") return "Últimos 6 meses";
  if (r.preset === "year") return "Este año";
  if (r.preset === "last-year") return "Año pasado";
  if (r.preset === "all") return "Todo el historial";
  // custom
  const { from, to } = resolveRange(r);
  if (from === to) return prettyDate(from);
  return `${prettyDate(from)} – ${prettyDate(to)}`;
}

function prettyDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-").map(Number);
  if (!y || !m || !d) return iso;
  const M = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  return `${d} ${M[m - 1]} ${y}`;
}

/* filtra txs según rango (los movimientos fuera del rango se omiten) */
function txsInRange(txs, range) {
  const { from, to } = resolveRange(range);
  return txs.filter((t) => t.date >= from && t.date <= to);
}

/* elige granularidad sensata según el ancho del rango */
function rangeGranularity(range) {
  const { from, to } = resolveRange(range);
  const days = Math.max(1, Math.round((new Date(to) - new Date(from)) / 86400000) + 1);
  if (days <= 14) return "day";
  if (days <= 92) return "week";
  if (days <= 730) return "month";
  return "month";
}

/* convierte una spec de gráfica del asistente en el payload que consume MiniChart */
function buildChart(spec, config, txs) {
  /* spec esperada del asistente:
     { kind: "line"|"bars"|"pie",
       title?: "string",
       groupBy: "day"|"week"|"month"|"category"|"account",
       metric: "income"|"expense"|"both"|"net"|"expense_by_category",
       dateFrom: "YYYY-MM-DD", dateTo: "YYYY-MM-DD",
       accountId?: "..." | null,
       includePassThrough?: false  // default false }
  */
  if (!spec || !spec.kind) return null;
  const kind = spec.kind;
  const includePT = !!spec.includePassThrough;
  const metric = spec.metric || "both";
  const groupBy = spec.groupBy || "month";
  // rango por defecto: mes actual
  let dateFrom = spec.dateFrom || today().slice(0, 7) + "-01";
  let dateTo = spec.dateTo || today();
  if (dateFrom > dateTo) { const tmp = dateFrom; dateFrom = dateTo; dateTo = tmp; }

  // filtrar
  let pool = txs.filter((t) => t.date >= dateFrom && t.date <= dateTo);
  if (spec.accountId) pool = pool.filter((t) => t.accountId === spec.accountId);
  // aplicar statTxs (excluye pass-through, agrega diferencias netas)
  pool = includePT ? pool : statTxs(pool).all;

  // helpers de buckets
  const buckets = [];
  if (groupBy === "day") {
    const d0 = new Date(dateFrom + "T12:00:00");
    const d1 = new Date(dateTo + "T12:00:00");
    for (let d = new Date(d0); d <= d1; d.setDate(d.getDate() + 1)) {
      const k = d.toISOString().slice(0, 10);
      buckets.push({ key: k, label: `${d.getDate()}` });
    }
  } else if (groupBy === "week") {
    // semanas lunes-domingo
    const d0 = new Date(dateFrom + "T12:00:00");
    const d1 = new Date(dateTo + "T12:00:00");
    const startMon = new Date(d0);
    const dow = startMon.getDay(); // 0 dom, 1 lun
    const offset = (dow === 0 ? -6 : 1 - dow);
    startMon.setDate(startMon.getDate() + offset);
    let n = 1;
    for (let s = new Date(startMon); s <= d1; s.setDate(s.getDate() + 7)) {
      const end = new Date(s); end.setDate(end.getDate() + 6);
      buckets.push({
        key: s.toISOString().slice(0, 10) + "..." + end.toISOString().slice(0, 10),
        label: `S${n}`,
        startISO: s.toISOString().slice(0, 10),
        endISO: end.toISOString().slice(0, 10),
      });
      n++;
    }
  } else if (groupBy === "month") {
    const [y0, m0] = dateFrom.split("-").map(Number);
    const [y1, m1] = dateTo.split("-").map(Number);
    let y = y0, m = m0;
    const MNAMES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
    while (y < y1 || (y === y1 && m <= m1)) {
      buckets.push({ key: `${y}-${String(m).padStart(2, "0")}`, label: MNAMES[m - 1] });
      m++; if (m > 12) { m = 1; y++; }
    }
  }

  function bucketOf(date) {
    if (groupBy === "day") return date;
    if (groupBy === "month") return date.slice(0, 7);
    if (groupBy === "week") {
      const b = buckets.find((b) => date >= b.startISO && date <= b.endISO);
      return b ? b.key : null;
    }
    return null;
  }

  // ===== gráfica de pastel: gastos por categoría =====
  if (kind === "pie" || metric === "expense_by_category") {
    const byCat = {};
    pool.filter((t) => t.type === "expense" && t.categoryId).forEach((t) => {
      byCat[t.categoryId] = (byCat[t.categoryId] || 0) + t.amount;
    });
    const segments = Object.entries(byCat)
      .map(([id, value]) => {
        const c = config.categories.find((x) => x.id === id);
        return { label: c ? `${c.emoji} ${c.name}` : "Sin categoría", value };
      })
      .sort((a, b) => b.value - a.value);
    return { kind: "pie", title: spec.title || "Gastos por categoría", segments };
  }

  // ===== series temporales (line/bars) =====
  if (groupBy === "category") {
    // barras por categoría
    const byCat = {};
    pool.filter((t) => t.type === (metric === "income" ? "income" : "expense") && t.categoryId).forEach((t) => {
      byCat[t.categoryId] = (byCat[t.categoryId] || 0) + t.amount;
    });
    const sorted = Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const xLabels = sorted.map(([id]) => {
      const c = config.categories.find((x) => x.id === id);
      return c ? c.name : "—";
    });
    const values = sorted.map(([, v]) => v);
    return {
      kind: kind === "line" ? "bars" : kind, // forzar barras para categorías
      title: spec.title || (metric === "income" ? "Ingresos por categoría" : "Gastos por categoría"),
      xLabels,
      series: [{ name: metric === "income" ? "Ingresos" : "Gastos", color: metric === "income" ? "#2E6F4E" : "#B8482A", values }],
    };
  }

  // por tiempo
  const incValues = buckets.map(() => 0);
  const expValues = buckets.map(() => 0);
  const netValues = buckets.map(() => 0);
  for (const t of pool) {
    const bk = bucketOf(t.date);
    if (!bk) continue;
    const idx = buckets.findIndex((b) => b.key === bk);
    if (idx < 0) continue;
    if (t.type === "income") { incValues[idx] += t.amount; netValues[idx] += t.amount; }
    else { expValues[idx] += t.amount; netValues[idx] -= t.amount; }
  }

  const xLabels = buckets.map((b) => b.label);
  let series = [];
  if (metric === "income") series = [{ name: "Ingresos", color: "#2E6F4E", values: incValues }];
  else if (metric === "expense") series = [{ name: "Gastos", color: "#B8482A", values: expValues }];
  else if (metric === "net") series = [{ name: "Flujo neto", color: "#B0863A", values: netValues }];
  else series = [
    { name: "Ingresos", color: "#2E6F4E", values: incValues },
    { name: "Gastos", color: "#B8482A", values: expValues },
  ];

  const titleMap = {
    day: "por día", week: "por semana", month: "por mes",
  };
  const defaultTitle = spec.title
    || (metric === "net" ? `Flujo neto ${titleMap[groupBy] || ""}` : `Ingresos vs gastos ${titleMap[groupBy] || ""}`);

  return { kind, title: defaultTitle.trim(), xLabels, series };
}

/* buscar categoría / cuenta por id o por nombre */
function findCat(config, ref, accountId) {
  if (ref == null) return null;
  const r = norm(String(ref)).trim();
  const byId = config.categories.find((c) => c.id === ref);
  if (byId) return byId;
  const pool = accountId ? config.categories.filter((c) => c.accountId === accountId) : config.categories;
  return pool.find((c) => norm(c.name).trim() === r)
    || pool.find((c) => norm(c.name).includes(r) && r.length >= 3) || null;
}
function findAcc(config, ref) {
  if (ref == null) return null;
  const r = norm(String(ref)).trim();
  return config.accounts.find((a) => a.id === ref)
    || config.accounts.find((a) => norm(a.name).trim() === r)
    || config.accounts.find((a) => norm(a.name).includes(r) && r.length >= 3) || null;
}

/* ========================== RECURRENCIAS ================================
   config.recurring = [
     {
       id, type:"income"|"expense", amount, description,
       categoryId, accountId,
       freq:"daily"|"weekly"|"biweekly"|"monthly"|"yearly",
       dayOfMonth?:number (para monthly/yearly, 1-31),
       month?:number (para yearly, 1-12),
       daysOfWeek?:[0..6] (para weekly: 0=domingo),
       startDate:"YYYY-MM-DD",
       endDate?:"YYYY-MM-DD",
       paused?:boolean,
       lastRunDate?:"YYYY-MM-DD" (última fecha cubierta — incluida)
     }
   ]
   Cada movimiento generado por una recurrencia tiene t.fromRecurring = <id> */

function dateKeyDaysAfter(iso, n) {
  const d = new Date(iso + "T12:00:00");
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}

/* todas las fechas en las que la recurrencia debe ejecutarse desde "fromDate" hasta "toDate" inclusivo */
function recurringDatesBetween(rule, fromDate, toDate) {
  const dates = [];
  if (!rule || !rule.startDate) return dates;
  const start = rule.startDate;
  const end = rule.endDate || toDate;
  // Empezar desde el día siguiente a lastRunDate, o desde startDate si nunca corrió
  let cursor = rule.lastRunDate ? dateKeyDaysAfter(rule.lastRunDate, 1) : start;
  if (cursor < start) cursor = start;
  const realEnd = end < toDate ? end : toDate;
  // safety: limita a 365 días por procesado para no colgarse
  let safety = 0;
  while (cursor <= realEnd && safety < 800) {
    const d = new Date(cursor + "T12:00:00");
    let matches = false;
    if (rule.freq === "daily") matches = true;
    else if (rule.freq === "weekly") {
      const dow = d.getDay();
      const days = Array.isArray(rule.daysOfWeek) && rule.daysOfWeek.length ? rule.daysOfWeek : [new Date(start + "T12:00:00").getDay()];
      matches = days.includes(dow);
    }
    else if (rule.freq === "biweekly") {
      const startD = new Date(start + "T12:00:00");
      const diff = Math.round((d - startD) / 86400000);
      matches = diff >= 0 && diff % 14 === 0;
    }
    else if (rule.freq === "monthly") {
      const dom = rule.dayOfMonth || new Date(start + "T12:00:00").getDate();
      // si dom > último día del mes, usar último día
      const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
      const targetDom = Math.min(dom, lastDay);
      matches = d.getDate() === targetDom;
    }
    else if (rule.freq === "yearly") {
      const startD = new Date(start + "T12:00:00");
      const mon = (rule.month || (startD.getMonth() + 1)) - 1;
      const dom = rule.dayOfMonth || startD.getDate();
      matches = d.getMonth() === mon && d.getDate() === dom;
    }
    if (matches) dates.push(cursor);
    cursor = dateKeyDaysAfter(cursor, 1);
    safety++;
  }
  return dates;
}

/* Procesar todas las recurrencias activas. Devuelve { config, txs, generated } */
function processRecurring(config, txs) {
  const rules = Array.isArray(config.recurring) ? config.recurring : [];
  if (!rules.length) return { config, txs, generated: 0 };
  const todayK = today();
  const newTxs = [];
  const updatedRules = rules.map((rule) => {
    if (rule.paused) return rule;
    if (!rule.amount || !rule.accountId) return rule;
    // si la categoría ya no existe o no pertenece a la cuenta, dejar null
    let categoryId = rule.categoryId;
    const cat = config.categories.find((c) => c.id === categoryId);
    if (!cat || cat.accountId !== rule.accountId) categoryId = null;

    const fires = recurringDatesBetween(rule, rule.startDate, todayK);
    if (!fires.length) return rule;
    fires.forEach((date) => {
      newTxs.push({
        id: uid(),
        type: rule.type || "expense",
        amount: Number(rule.amount) || 0,
        description: rule.description || "(recurrencia)",
        categoryId,
        accountId: rule.accountId,
        date,
        fromRecurring: rule.id,
      });
    });
    return { ...rule, lastRunDate: fires[fires.length - 1] };
  });
  if (!newTxs.length) return { config, txs, generated: 0 };
  return {
    config: { ...config, recurring: updatedRules },
    txs: [...newTxs, ...txs],
    generated: newTxs.length,
  };
}

/* aplicar las acciones que devuelve el asistente */
function applyActions(config0, txs0, actions) {
  let config = { ...config0, categories: [...config0.categories], accounts: [...config0.accounts] };
  let txs = [...txs0];
  const log = [];
  const ok = (text) => log.push({ ok: true, text });
  const no = (text) => log.push({ ok: false, text });

  for (const a of Array.isArray(actions) ? actions : []) {
    if (!a || !a.type) continue;
    try {
      if (a.type === "add_category") {
        const ac = findAcc(config, a.accountId ?? a.accountName) || config.accounts[0];
        const seed = SEED_KW[norm(a.name || "").trim()];
        config = { ...config, categories: [...config.categories, {
          id: uid(), name: a.name || "Categoría", emoji: a.emoji || "📦",
          type: a.categoryType === "income" ? "income" : "expense",
          accountId: ac ? ac.id : null, keywords: seed ? [...seed] : [],
        }] };
        ok(`Categoría creada en ${ac ? ac.name : "?"}: ${a.emoji || "📦"} ${a.name || ""}`);

      } else if (a.type === "remove_category") {
        const c = findCat(config, a.id ?? a.name, findAcc(config, a.accountId ?? a.accountName)?.id);
        if (c) { config = { ...config, categories: config.categories.filter((x) => x.id !== c.id) }; ok(`Categoría eliminada: ${c.name}`); }
        else no(`No encontré la categoría «${a.name ?? a.id}»`);

      } else if (a.type === "edit_category") {
        const c = findCat(config, a.id ?? a.name, findAcc(config, a.accountId ?? a.accountName)?.id);
        if (c) {
          const u = { ...c };
          if (a.newName) u.name = a.newName;
          if (a.newEmoji) u.emoji = a.newEmoji;
          if (a.newType) u.type = a.newType === "income" ? "income" : "expense";
          config = { ...config, categories: config.categories.map((x) => x.id === c.id ? u : x) };
          ok(`Categoría actualizada: ${u.emoji} ${u.name}`);
        } else no(`No encontré la categoría «${a.name ?? a.id}»`);

      } else if (a.type === "set_account_categories") {
        const ac = findAcc(config, a.accountId ?? a.accountName);
        const incoming = Array.isArray(a.categories) ? a.categories : [];
        if (!ac) no(`No encontré la cuenta «${a.accountName ?? a.accountId}»`);
        else if (!incoming.length) no(`No me diste las categorías para «${ac.name}»`);
        else {
          const ofAcc = config.categories.filter((c) => c.accountId === ac.id);
          const oldExpense = ofAcc.filter((c) => c.type === "expense");
          const keepIncome = ofAcc.filter((c) => c.type === "income");
          const newExpense = incoming.map((it) => {
            const nm = (it.name || "Categoría").trim();
            const prev = oldExpense.find((c) => norm(c.name).trim() === norm(nm).trim());
            const seed = SEED_KW[norm(nm).trim()];
            return {
              id: uid(), name: nm, emoji: it.emoji || (prev ? prev.emoji : "📦"),
              type: "expense", accountId: ac.id,
              keywords: prev ? [...(prev.keywords || [])] : (seed ? [...seed] : []),
            };
          });
          // remapear movimientos por nombre para no perder historial
          const nameToNew = {};
          newExpense.forEach((c) => { nameToNew[norm(c.name).trim()] = c.id; });
          txs = txs.map((t) => {
            const oc = oldExpense.find((c) => c.id === t.categoryId);
            if (!oc) return t;
            const nid = nameToNew[norm(oc.name).trim()];
            return { ...t, categoryId: nid || null };
          });
          const others = config.categories.filter((c) => c.accountId !== ac.id);
          config = { ...config, categories: [...others, ...keepIncome, ...newExpense] };
          ok(`Categorías de gasto de ${ac.name} actualizadas (${newExpense.length})`);
        }

      } else if (a.type === "add_account") {
        const nid = uid();
        config = {
          ...config,
          accounts: [...config.accounts, { id: nid, name: a.name || "Cuenta", initialBalance: Number(a.initialBalance) || 0 }],
          categories: [...config.categories, ...defaultCatsForAccount(nid)],
          accountMode: "multiple",
        };
        ok(`Cuenta creada: ${a.name || ""} (saldo inicial ${fmt(Number(a.initialBalance) || 0)})`);

      } else if (a.type === "remove_account") {
        const ac = findAcc(config, a.id ?? a.name);
        if (!ac) no(`No encontré la cuenta «${a.name ?? a.id}»`);
        else if (config.accounts.length <= 1) no(`No puedo eliminar la única cuenta.`);
        else {
          config = {
            ...config,
            accounts: config.accounts.filter((x) => x.id !== ac.id),
            categories: config.categories.filter((c) => c.accountId !== ac.id),
          };
          config.accountMode = config.accounts.length > 1 ? "multiple" : "single";
          ok(`Cuenta eliminada: ${ac.name}`);
        }

      } else if (a.type === "edit_account") {
        const ac = findAcc(config, a.id ?? a.name);
        if (ac) {
          const u = { ...ac };
          if (a.newName) u.name = a.newName;
          if (a.newInitialBalance != null) u.initialBalance = Number(a.newInitialBalance) || 0;
          config = { ...config, accounts: config.accounts.map((x) => x.id === ac.id ? u : x) };
          ok(`Cuenta actualizada: ${u.name}`);
        } else no(`No encontré la cuenta «${a.name ?? a.id}»`);

      } else if (a.type === "add_transaction") {
        const ac = findAcc(config, a.accountName) || config.accounts[0];
        const c = findCat(config, a.categoryName, ac ? ac.id : null);
        const tx = {
          id: uid(), type: a.txType === "income" ? "income" : "expense",
          amount: Math.abs(Number(a.amount) || 0), description: a.description || "",
          categoryId: c ? c.id : null, accountId: ac ? ac.id : null, date: a.date || today(),
          passThrough: !!a.passThrough, linkedTxId: a.linkedTxId || null,
        };
        txs = [tx, ...txs];
        ok(`${tx.type === "income" ? "Ingreso" : "Gasto"}${tx.passThrough ? " (de paso)" : ""} registrado: ${tx.description || (c ? c.name : "")} · ${fmt(tx.amount)}`);

      } else if (a.type === "delete_transaction") {
        const t = txs.find((x) => x.id === a.id);
        if (t) { txs = txs.filter((x) => x.id !== a.id); ok(`Movimiento eliminado: ${t.description || fmt(t.amount)}`); }
        else no(`No encontré ese movimiento`);

      } else if (a.type === "edit_transaction") {
        const t = txs.find((x) => x.id === a.id);
        if (t) {
          const u = { ...t };
          if (a.amount != null) u.amount = Math.abs(Number(a.amount) || 0);
          if (a.description != null) u.description = a.description;
          if (a.txType) u.type = a.txType === "income" ? "income" : "expense";
          if (a.date) u.date = a.date;
          if (a.accountName) { const ac = findAcc(config, a.accountName); if (ac) u.accountId = ac.id; }
          if (a.categoryName) { const c = findCat(config, a.categoryName, u.accountId); if (c) u.categoryId = c.id; }
          if (a.passThrough != null) u.passThrough = !!a.passThrough;
          if (a.linkedTxId !== undefined) u.linkedTxId = a.linkedTxId || null;
          txs = txs.map((x) => x.id === t.id ? u : x);
          ok(`Movimiento actualizado: ${u.description || fmt(u.amount)}`);
        } else no(`No encontré ese movimiento`);

      } else if (a.type === "mark_passthrough") {
        // marca uno o dos movimientos como "no real" y opcionalmente los vincula
        const ids = Array.isArray(a.ids) ? a.ids : (a.id ? [a.id] : []);
        const found = ids.map((id) => txs.find((x) => x.id === id)).filter(Boolean);
        if (!found.length) { no(`No encontré los movimientos a marcar como dinero de paso`); }
        else {
          // si son exactamente 2 y son de tipos opuestos, vincularlos
          let linkA = null, linkB = null;
          if (found.length === 2 && found[0].type !== found[1].type) {
            linkA = found[0].id; linkB = found[1].id;
          }
          const updates = new Map();
          found.forEach((t) => {
            const u = { ...t, passThrough: true };
            if (linkA && linkB) u.linkedTxId = t.id === linkA ? linkB : linkA;
            updates.set(t.id, u);
          });
          txs = txs.map((x) => updates.has(x.id) ? updates.get(x.id) : x);
          if (linkA && linkB) {
            const inc = found.find((t) => t.type === "income");
            const exp = found.find((t) => t.type === "expense");
            const diff = (inc?.amount || 0) - (exp?.amount || 0);
            const note = Math.abs(diff) < 0.01 ? "Sin diferencia neta"
              : diff > 0 ? `Ganancia neta: ${fmt(diff)}` : `Pusiste de tu bolsa: ${fmt(Math.abs(diff))}`;
            ok(`Vinculados como reembolso: ${inc?.description || fmt(inc?.amount || 0)} ↔ ${exp?.description || fmt(exp?.amount || 0)}. ${note}`);
          } else {
            found.forEach((t) => ok(`Marcado como dinero de paso: ${t.description || fmt(t.amount)}`));
          }
        }

      } else if (a.type === "unmark_passthrough") {
        const ids = Array.isArray(a.ids) ? a.ids : (a.id ? [a.id] : []);
        const targets = ids.map((id) => txs.find((x) => x.id === id)).filter(Boolean);
        if (!targets.length) { no(`No encontré los movimientos a desmarcar`); }
        else {
          const setIds = new Set();
          targets.forEach((t) => {
            setIds.add(t.id);
            if (t.linkedTxId) setIds.add(t.linkedTxId);
          });
          txs = txs.map((x) => setIds.has(x.id) ? { ...x, passThrough: false, linkedTxId: null } : x);
          ok(`Desmarcados ${setIds.size} movimiento${setIds.size === 1 ? "" : "s"} (vuelven a contar como reales)`);
        }

      /* ============== RECURRENCIAS ============== */
      } else if (a.type === "add_recurring") {
        const ac = findAcc(config, a.accountId ?? a.accountName);
        if (!ac) { no(`No encontré la cuenta «${a.accountName ?? a.accountId}»`); }
        else {
          const cat = a.categoryName || a.categoryId
            ? findCat(config, a.categoryId ?? a.categoryName, ac.id)
            : null;
          const validFreqs = ["daily","weekly","biweekly","monthly","yearly"];
          const freq = validFreqs.includes(a.freq) ? a.freq : "monthly";
          const rule = {
            id: uid(),
            type: a.txType === "income" ? "income" : "expense",
            amount: Number(a.amount) || 0,
            description: a.description || "",
            categoryId: cat ? cat.id : null,
            accountId: ac.id,
            freq,
            dayOfMonth: a.dayOfMonth != null ? Number(a.dayOfMonth) : undefined,
            month: a.month != null ? Number(a.month) : undefined,
            daysOfWeek: Array.isArray(a.daysOfWeek) ? a.daysOfWeek.map(Number) : undefined,
            startDate: a.startDate || today(),
            endDate: a.endDate || undefined,
            paused: false,
            active: true,
          };
          config = { ...config, recurring: [...(config.recurring || []), rule] };
          // ejecutar de inmediato lo pendiente
          const r = processRecurring(config, txs);
          config = r.config; txs = r.txs;
          const freqLabel = { daily: "diaria", weekly: "semanal", biweekly: "quincenal", monthly: "mensual", yearly: "anual" }[freq];
          ok(`Recurrencia ${freqLabel} creada: ${rule.type === "income" ? "+" : "−"}$${rule.amount} en ${ac.name}${r.generated > 0 ? ` (se aplicaron ${r.generated} de inmediato)` : ""}`);
        }

      } else if (a.type === "edit_recurring") {
        const rules = config.recurring || [];
        const rule = rules.find((r) => r.id === a.id);
        if (!rule) { no(`No encontré la recurrencia con id ${a.id}`); }
        else {
          const u = { ...rule };
          if (a.amount != null) u.amount = Number(a.amount);
          if (a.description != null) u.description = a.description;
          if (a.txType) u.type = a.txType === "income" ? "income" : "expense";
          if (a.freq) u.freq = a.freq;
          if (a.dayOfMonth != null) u.dayOfMonth = Number(a.dayOfMonth);
          if (a.month != null) u.month = Number(a.month);
          if (Array.isArray(a.daysOfWeek)) u.daysOfWeek = a.daysOfWeek.map(Number);
          if (a.startDate) u.startDate = a.startDate;
          if (a.endDate !== undefined) u.endDate = a.endDate || undefined;
          if (a.accountId || a.accountName) {
            const ac = findAcc(config, a.accountId ?? a.accountName);
            if (ac) u.accountId = ac.id;
          }
          if (a.categoryId || a.categoryName) {
            const cat = findCat(config, a.categoryId ?? a.categoryName, u.accountId);
            if (cat) u.categoryId = cat.id;
          }
          config = { ...config, recurring: rules.map((r) => r.id === u.id ? u : r) };
          ok(`Recurrencia actualizada`);
        }

      } else if (a.type === "delete_recurring") {
        const rules = config.recurring || [];
        const before = rules.length;
        const newRules = rules.filter((r) => r.id !== a.id);
        if (newRules.length === before) { no(`No encontré esa recurrencia`); }
        else {
          config = { ...config, recurring: newRules };
          // borrar también los movimientos ya generados, si el usuario lo pide explícito
          if (a.alsoDeleteGenerated) {
            const removed = txs.filter((t) => t.fromRecurring === a.id).length;
            txs = txs.filter((t) => t.fromRecurring !== a.id);
            ok(`Recurrencia eliminada y ${removed} movimientos generados también borrados`);
          } else {
            ok(`Recurrencia eliminada (los movimientos ya generados se mantienen)`);
          }
        }

      } else if (a.type === "pause_recurring") {
        const rules = config.recurring || [];
        const rule = rules.find((r) => r.id === a.id);
        if (!rule) { no(`No encontré esa recurrencia`); }
        else {
          config = { ...config, recurring: rules.map((r) => r.id === a.id ? { ...r, paused: true } : r) };
          ok(`Recurrencia pausada`);
        }

      } else if (a.type === "resume_recurring") {
        const rules = config.recurring || [];
        const rule = rules.find((r) => r.id === a.id);
        if (!rule) { no(`No encontré esa recurrencia`); }
        else {
          config = { ...config, recurring: rules.map((r) => r.id === a.id ? { ...r, paused: false } : r) };
          // procesar lo pendiente
          const r2 = processRecurring(config, txs);
          config = r2.config; txs = r2.txs;
          ok(`Recurrencia reanudada${r2.generated > 0 ? ` (se aplicaron ${r2.generated} movimientos pendientes)` : ""}`);
        }

      /* ============== ACCIONES MASIVAS (BULK) ============== */
      } else if (a.type === "bulk_edit_category") {
        // re-categoriza muchos movimientos. Filtros opcionales: ids, fromCategoryId, accountId, keyword, fromDate, toDate
        const ac = a.accountId || a.accountName ? findAcc(config, a.accountId ?? a.accountName) : null;
        const targetCat = findCat(config, a.toCategoryId ?? a.toCategoryName, ac?.id);
        if (!targetCat) { no(`No encontré la categoría destino «${a.toCategoryName ?? a.toCategoryId}»`); }
        else {
          const fromCat = a.fromCategoryId || a.fromCategoryName ? findCat(config, a.fromCategoryId ?? a.fromCategoryName, ac?.id) : null;
          const ids = Array.isArray(a.ids) ? new Set(a.ids) : null;
          const kw = a.keyword ? norm(a.keyword) : null;
          let touched = 0;
          txs = txs.map((t) => {
            if (ids && !ids.has(t.id)) return t;
            if (!ids && ac && t.accountId !== ac.id) return t;
            if (!ids && fromCat && t.categoryId !== fromCat.id) return t;
            if (!ids && kw && !norm(t.description || "").includes(kw)) return t;
            if (!ids && a.fromDate && t.date < a.fromDate) return t;
            if (!ids && a.toDate && t.date > a.toDate) return t;
            if (t.categoryId === targetCat.id) return t;
            touched++;
            return { ...t, categoryId: targetCat.id };
          });
          if (touched > 0) ok(`Re-categorizados ${touched} movimientos a ${targetCat.emoji} ${targetCat.name}`);
          else no(`No encontré movimientos que coincidan con los filtros`);
        }

      } else if (a.type === "bulk_delete") {
        const ids = Array.isArray(a.ids) ? new Set(a.ids) : null;
        const ac = a.accountId || a.accountName ? findAcc(config, a.accountId ?? a.accountName) : null;
        const fromCat = a.categoryId || a.categoryName ? findCat(config, a.categoryId ?? a.categoryName, ac?.id) : null;
        const kw = a.keyword ? norm(a.keyword) : null;
        const before = txs.length;
        txs = txs.filter((t) => {
          if (ids) return !ids.has(t.id);
          // criterios combinables
          let match = true;
          if (ac && t.accountId !== ac.id) match = false;
          if (fromCat && t.categoryId !== fromCat.id) match = false;
          if (kw && !norm(t.description || "").includes(kw)) match = false;
          if (a.fromDate && t.date < a.fromDate) match = false;
          if (a.toDate && t.date > a.toDate) match = false;
          if (a.txType && t.type !== a.txType) match = false;
          return !match;  // si NO matchea todos los filtros, se conserva
        });
        const removed = before - txs.length;
        if (removed > 0) ok(`Eliminados ${removed} movimientos`);
        else no(`No encontré movimientos que coincidan con los filtros`);

      } else if (a.type === "bulk_move_account") {
        const target = findAcc(config, a.toAccountId ?? a.toAccountName);
        if (!target) { no(`No encontré la cuenta destino`); }
        else {
          const ids = Array.isArray(a.ids) ? new Set(a.ids) : null;
          const fromAc = a.fromAccountId || a.fromAccountName ? findAcc(config, a.fromAccountId ?? a.fromAccountName) : null;
          const kw = a.keyword ? norm(a.keyword) : null;
          let touched = 0;
          txs = txs.map((t) => {
            let match = true;
            if (ids) match = ids.has(t.id);
            else {
              if (fromAc && t.accountId !== fromAc.id) match = false;
              if (kw && !norm(t.description || "").includes(kw)) match = false;
              if (a.fromDate && t.date < a.fromDate) match = false;
              if (a.toDate && t.date > a.toDate) match = false;
            }
            if (!match) return t;
            if (t.accountId === target.id) return t;
            touched++;
            // al mover de cuenta, la categoría ya no aplica (pertenece a otra cuenta)
            return { ...t, accountId: target.id, categoryId: null };
          });
          if (touched > 0) ok(`Movidos ${touched} movimientos a ${target.name} (su categoría se reinició)`);
          else no(`No encontré movimientos que coincidan`);
        }

      /* ============== CONFIGURACIÓN ============== */
      } else if (a.type === "set_date_range") {
        const validPresets = ["today","week","month","last-month","3m","6m","year","last-year","all","custom"];
        const preset = validPresets.includes(a.preset) ? a.preset : "month";
        const newRange = preset === "custom"
          ? { preset: "custom", from: a.from || today(), to: a.to || today() }
          : { preset };
        config = { ...config, dateRange: newRange };
        ok(`Rango cambiado a: ${rangeLabel(newRange)}`);

      } else {
        no(`Acción no reconocida: ${a.type}`);
      }
    } catch (e) {
      no(`Hubo un error aplicando un cambio.`);
    }
  }
  return { config, txs, log };
}

/* system prompt del asistente, con el estado actual de la app */
function assistantSystem(config, txs) {
  const cuentas = config.accounts.map((a) => ({
    id: a.id, nombre: a.name, saldoInicial: a.initialBalance || 0, saldoActual: accountBalance(config, txs, a.id),
    categorias: config.categories.filter((c) => c.accountId === a.id)
      .map((c) => ({ id: c.id, nombre: c.name, emoji: c.emoji, tipo: c.type })),
  }));
  // hasta 100 movimientos recientes (de los más nuevos)
  const sorted = [...txs].sort((a, b) => b.date.localeCompare(a.date) || (b.id || "").localeCompare(a.id || ""));
  const recent = sorted.slice(0, 100).map((t) => ({
    id: t.id, tipo: t.type, monto: t.amount, concepto: t.description,
    categoria: (config.categories.find((c) => c.id === t.categoryId) || {}).name || null,
    cuenta: (config.accounts.find((a) => a.id === t.accountId) || {}).name || null, fecha: t.date,
    dineroDePaso: !!t.passThrough,
    vinculadoCon: t.linkedTxId || null,
    deRecurrencia: t.fromRecurring || null,
  }));

  /* ============ AGREGADOS PRE-CALCULADOS ============ */
  // rango activo del usuario
  const range = config.dateRange || DEFAULT_RANGE;
  const { from: rangeFrom, to: rangeTo } = resolveRange(range);
  const rangeTxs = txsInRange(txs, range);
  const rangeStat = statTxs(rangeTxs).all;
  const rangeIncome = rangeStat.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const rangeExpense = rangeStat.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  // top categorías de gasto e ingreso en el rango
  const expByCat = {}, incByCat = {};
  rangeStat.forEach((t) => {
    if (!t.categoryId) return;
    if (t.type === "expense") expByCat[t.categoryId] = (expByCat[t.categoryId] || 0) + t.amount;
    else incByCat[t.categoryId] = (incByCat[t.categoryId] || 0) + t.amount;
  });
  const topRows = (obj) => Object.entries(obj)
    .map(([id, amt]) => {
      const c = config.categories.find((x) => x.id === id);
      return c ? { categoria: c.name, monto: amt } : null;
    }).filter(Boolean).sort((a, b) => b.monto - a.monto).slice(0, 8);

  // saldo neto histórico
  const saldoNeto = config.accounts.reduce((s, a) => s + accountBalance(config, txs, a.id), 0);

  // por mes (últimos 6 meses) para comparaciones
  const monthly = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date();
    d.setMonth(d.getMonth() - i);
    const mk = d.toISOString().slice(0, 7);
    const monthTxs = txs.filter((t) => t.date.slice(0, 7) === mk);
    const ms = statTxs(monthTxs).all;
    monthly.push({
      mes: mk,
      ingresos: ms.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0),
      gastos: ms.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0),
      movimientos: monthTxs.length,
    });
  }

  const agregados = {
    rango: { preset: range.preset, desde: rangeFrom, hasta: rangeTo, etiqueta: rangeLabel(range) },
    enRango: { ingresos: rangeIncome, gastos: rangeExpense, flujoNeto: rangeIncome - rangeExpense, movimientos: rangeTxs.length },
    topGastosEnRango: topRows(expByCat),
    topIngresosEnRango: topRows(incByCat),
    porMes6m: monthly,
    saldoNetoGlobal: saldoNeto,
  };

  // recurrencias actuales
  const recurrencias = (config.recurring || []).map((r) => ({
    id: r.id,
    tipo: r.type,
    monto: r.amount,
    concepto: r.description,
    cuenta: (config.accounts.find((a) => a.id === r.accountId) || {}).name || null,
    categoria: (config.categories.find((c) => c.id === r.categoryId) || {}).name || null,
    frecuencia: r.freq,
    diaDelMes: r.dayOfMonth || null,
    diasDeSemana: r.daysOfWeek || null,
    desde: r.startDate,
    hasta: r.endDate || null,
    pausada: !!r.paused,
    ultimaCorrida: r.lastRunDate || null,
  }));

  return `Eres el asistente de "Zafi", una app de finanzas personales con IA. El usuario habla español de México. Te pide cosas en lenguaje natural y tú las ejecutas o respondes sus dudas.

IMPORTANTE: cada CUENTA tiene SU PROPIA lista de categorías. Las categorías NO son globales; pertenecen a una cuenta.

ESTADO ACTUAL DE LA APP:
Cuentas (cada una con sus categorías): ${JSON.stringify(cuentas)}
Movimientos recientes (últimos 100): ${JSON.stringify(recent)}
Recurrencias activas: ${JSON.stringify(recurrencias)}
Agregados pre-calculados: ${JSON.stringify(agregados)}
Fecha de hoy: ${today()}

RESPONDE SIEMPRE con UN SOLO objeto JSON válido, sin markdown ni texto fuera del JSON:
{"message":"texto breve para el usuario","actions":[ ... ]}

ACCIONES disponibles (usa "actions":[] si solo respondes una pregunta):
- {"type":"set_account_categories","accountId":"<id de la cuenta>","categories":[{"name":"...","emoji":"<emoji>"},...]}  → REEMPLAZA todas las categorías de GASTO de esa cuenta por la lista dada. Úsala cuando el usuario pida "cambia/pon las categorías de <cuenta> a esto: ...". Ponle un emoji apropiado a cada una. Las categorías de ingreso de esa cuenta NO se tocan.
- {"type":"add_category","accountId":"<id de la cuenta>","name":"...","emoji":"<emoji>","categoryType":"income"|"expense"}
- {"type":"edit_category","id":"<id de la categoría>","newName":"...","newEmoji":"...","newType":"income"|"expense"}
- {"type":"remove_category","id":"<id de la categoría>"}
- {"type":"add_account","name":"...","initialBalance":<número>}
- {"type":"edit_account","id":"<id de la cuenta>","newName":"...","newInitialBalance":<número>}
- {"type":"remove_account","id":"<id de la cuenta>"}
- {"type":"add_transaction","txType":"income"|"expense","amount":<número>,"description":"...","categoryName":"<categoría existente de esa cuenta>","accountName":"<nombre de cuenta>","date":"YYYY-MM-DD","passThrough":<true|false opcional>}
- {"type":"edit_transaction","id":"<id del movimiento>","amount":<número>,"description":"...","categoryName":"...","accountName":"...","txType":"...","date":"...","passThrough":<true|false>,"linkedTxId":"<id o null>"}
- {"type":"delete_transaction","id":"<id del movimiento>"}
- {"type":"show_chart","spec":{ ... }}  → muestra una gráfica inline en el chat. La spec define qué graficar (ver sección 📊 abajo). Úsala cuando el usuario te PIDA una gráfica.

🔁 RECURRENCIAS (movimientos automáticos):
- {"type":"add_recurring","txType":"income"|"expense","amount":<n>,"description":"...","accountName":"...","categoryName":"...","freq":"daily"|"weekly"|"biweekly"|"monthly"|"yearly","dayOfMonth":<1-31 opc>,"month":<1-12 opc>,"daysOfWeek":[<0-6> opc],"startDate":"YYYY-MM-DD opc","endDate":"YYYY-MM-DD opc"}
   Ejemplos para freq:
   • "daily" → todos los días desde startDate (default: hoy)
   • "weekly" + daysOfWeek:[1] → cada lunes (0=domingo, 1=lunes, ..., 6=sábado)
   • "biweekly" → cada 14 días desde startDate
   • "monthly" + dayOfMonth:15 → el día 15 de cada mes (si el mes no llega a ese día, se aplica el último)
   • "yearly" + month:12 + dayOfMonth:25 → cada 25 de diciembre
   Cuando creas una recurrencia, los movimientos pendientes desde startDate hasta hoy se aplican automáticamente.
- {"type":"edit_recurring","id":"<id>","amount":...,"description":"...","freq":"...","dayOfMonth":...,"endDate":"..."} — solo incluye los campos que cambian
- {"type":"delete_recurring","id":"<id>","alsoDeleteGenerated":<true|false opc>}  → si alsoDeleteGenerated=true, también borra todos los movimientos ya creados por esa recurrencia
- {"type":"pause_recurring","id":"<id>"}  → la regla deja de ejecutarse hasta que la reanuden
- {"type":"resume_recurring","id":"<id>"}  → al reanudar, aplica todos los movimientos pendientes desde lastRunDate hasta hoy

📦 ACCIONES MASIVAS (BULK):
- {"type":"bulk_edit_category","toCategoryName":"<categoría destino>","accountName":"<cuenta opc>","fromCategoryName":"<cuenta origen opc>","keyword":"<filtro opc>","fromDate":"YYYY-MM-DD opc","toDate":"YYYY-MM-DD opc","ids":["<id1>",...] (opc)}
   Re-categoriza varios movimientos al mismo tiempo. Filtra por ids, o por cualquier combinación de: accountName + fromCategoryName + keyword (en concepto) + rango de fechas.
   Ejemplos: "marca todo lo que diga OXXO como categoría Tiendas" → keyword:"OXXO", toCategoryName:"Tiendas".
- {"type":"bulk_delete","ids":["<id1>",...] (opc),"accountName":"...","categoryName":"...","keyword":"...","txType":"income"|"expense","fromDate":"...","toDate":"..."}
   Elimina varios movimientos. Filtra por ids, o por combinación de filtros (todos opcionales pero AL MENOS uno requerido).
   Ejemplos: "borra todos los Spotify" → keyword:"Spotify". "borra los gastos del 15 de mayo" → fromDate:"2026-05-15", toDate:"2026-05-15", txType:"expense".
- {"type":"bulk_move_account","toAccountName":"<cuenta destino>","fromAccountName":"<opc>","keyword":"<opc>","fromDate":"...","toDate":"...","ids":[...]}
   Mueve movimientos a otra cuenta. NOTA: al cambiar de cuenta la categoría se reinicia (porque pertenecía a otra cuenta).

⚙️ CONFIGURACIÓN:
- {"type":"set_date_range","preset":"today"|"week"|"month"|"last-month"|"3m"|"6m"|"year"|"last-year"|"all"|"custom","from":"YYYY-MM-DD (custom)","to":"YYYY-MM-DD (custom)"}
   Cambia el rango global de toda la app (afecta Inicio, Movimientos, Categorías y Estadísticas).

REGLAS:
- Para editar/eliminar categorías, cuentas o movimientos usa SIEMPRE el "id" exacto del estado de arriba.
- Para set_account_categories y add_category usa el "id" de la cuenta correcta.
- Si el usuario lista varias categorías para una cuenta, asume que son categorías de GASTO salvo que diga lo contrario.
- Si pide cambiar categorías de varias cuentas, devuelve una acción set_account_categories por cada cuenta.
- Para add_transaction, "categoryName" debe existir DENTRO de la cuenta indicada.
- Cada categoría nueva necesita un emoji apropiado.
- En edit_* incluye solo los campos que cambian.
- "message" debe ser corto y confirmar en español lo que hiciste.
- Si te piden algo que no puedes hacer, dilo en "message" sin inventar acciones.
- Si la petición es ambigua, pregunta en "message" y NO ejecutes acciones todavía.
- Para preguntas sobre gastos/ingresos usa los datos del estado de arriba.

SALDO INICIAL DE UNA CUENTA — IMPORTANTE:
El "saldoInicial" de cada cuenta es el monto base que tenía la cuenta antes de empezar a registrar movimientos. Es el campo "initialBalance" del modelo y se actualiza con la acción "edit_account" pasando "newInitialBalance".
Cuando el usuario diga cosas como:
- "ponle saldo inicial a Santander de 19039"
- "el saldo inicial de BBVA es 5000"
- "Santander empezó con 19039"
- "saldo inicial de Santander mayo 19039" (interpreta el mes como contexto, no como campo separado)
…SIEMPRE usa la acción {"type":"edit_account","id":"<id de la cuenta>","newInitialBalance":<número>}. Quita comas y signos del número: "19,039" → 19039.

🧠 INTELIGENCIA ANALÍTICA:
Tienes pre-calculado en "Agregados pre-calculados" todo lo necesario para responder preguntas analíticas SIN inventar números:
- agregados.enRango → ingresos, gastos, flujo neto del rango activo
- agregados.topGastosEnRango y topIngresosEnRango → categorías ordenadas con monto
- agregados.porMes6m → últimos 6 meses (cada uno con ingresos, gastos, movimientos) para comparaciones
- agregados.saldoNetoGlobal → patrimonio neto sumando todas las cuentas

Cuando el usuario pregunte cosas como:
- "¿cómo voy este mes?" → resume con agregados.enRango y top categorías
- "compara mayo con abril" → usa porMes6m, calcula la diferencia porcentual
- "¿en qué gasto más?" → top de topGastosEnRango
- "¿cuánto he gastado en Uber?" → busca en "Movimientos recientes" los que tienen "uber" en concepto y suma. Si necesitas mostrarlos, NO uses una acción — solo describe en "message". Si quieres mostrarlos visualmente, usa show_chart.
- "dame mis 10 gastos más grandes" → de "Movimientos recientes" filtra tipo expense, ordena por monto desc, toma los primeros 10, descríbelos
- "¿cuándo me toca el sueldo?" → revisa "Recurrencias activas" y calcula la próxima fecha

Siempre que respondas con números, redondea a pesos enteros (sin decimales) salvo que el usuario los pida explícitos. Formato: $1,234 con coma de miles.

CAPTURA MÚLTIPLE EN UN MENSAJE:
Si el usuario dice varios movimientos en una sola frase ("pagué 250 oxxo, recibí 5000, compré 1200 walmart"), devuelve UN array de varias acciones add_transaction. NO le preguntes uno por uno; intenta inferir todo y al final del "message" pregunta si algo quedó mal.

EJEMPLOS de respuesta correcta:

📊 GRÁFICAS — show_chart:
Cuando el usuario pida ver una gráfica, devuelve una acción show_chart con una "spec" que la app convierte en gráfica inline. NO calcules los datos tú, solo describe qué quieres graficar.

Formato de spec:
{
  "kind": "line" | "bars" | "pie",
  "title": "Texto del título (opcional)",
  "groupBy": "day" | "week" | "month" | "category" | "account",
  "metric": "income" | "expense" | "both" | "net" | "expense_by_category",
  "dateFrom": "YYYY-MM-DD",
  "dateTo": "YYYY-MM-DD",
  "accountId": "<id de cuenta o null para todas>",
  "includePassThrough": false
}

GUÍA PARA ELEGIR kind:
- "line" → tendencias en el tiempo (cómo evoluciona algo día/semana/mes)
- "bars" → comparar periodos o categorías lado a lado
- "pie" → distribución (gastos por categoría)

GUÍA PARA metric:
- "income" → solo ingresos
- "expense" → solo gastos
- "both" → dos series, ingresos y gastos por separado (lo más común para "ingresos vs gastos")
- "net" → flujo neto (ingresos − gastos por periodo)
- "expense_by_category" → gastos agrupados por categoría (se fuerza kind=pie)

GUÍA PARA fechas:
- "este mes" → dateFrom = primer día del mes actual, dateTo = hoy (${today()})
- "mayo 2026" → dateFrom = "2026-05-01", dateTo = "2026-05-31"
- "últimos 3 meses" → dateFrom = primer día de hace 3 meses, dateTo = hoy
- "este año" → dateFrom = "${today().slice(0, 4)}-01-01", dateTo = hoy
- "semana pasada" / "esta semana" → calcula el rango lunes-domingo correspondiente

GUÍA PARA groupBy:
- "day" → para rangos cortos (una semana o quincena)
- "week" → para 1 a 3 meses
- "month" → para 3 a 12 meses
- "category" → cuando preguntan "por categoría" sin pastel
- "account" → cuando comparan cuentas

REGLAS DE GRÁFICAS:
- Si el usuario menciona una sola cuenta ("solo de BBVA", "en Santander"), pon "accountId" con el id correcto.
- Si no especifica cuenta, deja accountId en null para usar todas.
- Por defecto, NO incluyas pass-through (includePassThrough: false) para que los reembolsos no inflen las cifras.
- En "message" describe brevemente qué muestra la gráfica.
- Puedes devolver más de una gráfica en una sola respuesta (varias show_chart en actions).

EJEMPLOS DE GRÁFICAS:

Usuario: "muéstrame una gráfica semana a semana del mes de mayo, ingresos vs gastos de líneas"
Tu respuesta:
{"message":"Aquí tienes mayo por semana 👇","actions":[{"type":"show_chart","spec":{"kind":"line","title":"Mayo · ingresos vs gastos","groupBy":"week","metric":"both","dateFrom":"${today().slice(0,4)}-05-01","dateTo":"${today().slice(0,4)}-05-31"}}]}

Usuario: "pastel de mis gastos del mes"
Tu respuesta:
{"message":"Tus gastos del mes por categoría:","actions":[{"type":"show_chart","spec":{"kind":"pie","metric":"expense_by_category","dateFrom":"${today().slice(0,7)}-01","dateTo":"${today()}"}}]}

Usuario: "barras de gastos por categoría este mes"
Tu respuesta:
{"message":"Aquí tienes:","actions":[{"type":"show_chart","spec":{"kind":"bars","groupBy":"category","metric":"expense","dateFrom":"${today().slice(0,7)}-01","dateTo":"${today()}"}}]}

Usuario: "cómo van mis ingresos los últimos 6 meses"
Tu respuesta (con cálculo de rango: hoy es ${today()}, hace 6 meses sería el primer día del mes que es 5 meses atrás):
{"message":"Tus ingresos de los últimos 6 meses:","actions":[{"type":"show_chart","spec":{"kind":"line","groupBy":"month","metric":"income","dateFrom":"<calcula>","dateTo":"${today()}"}}]}

Usuario: "compara BBVA y Santander este mes"
Cuando el usuario quiere comparar 2 cuentas, devuelve DOS gráficas, una por cuenta:
{"message":"Aquí los tienes lado a lado:","actions":[
  {"type":"show_chart","spec":{"kind":"bars","title":"BBVA","groupBy":"week","metric":"both","accountId":"<id BBVA>","dateFrom":"${today().slice(0,7)}-01","dateTo":"${today()}"}},
  {"type":"show_chart","spec":{"kind":"bars","title":"Santander","groupBy":"week","metric":"both","accountId":"<id Santander>","dateFrom":"${today().slice(0,7)}-01","dateTo":"${today()}"}}
]}

EJEMPLOS DE OTRAS ACCIONES:

Usuario: "ponle saldo inicial a Santander de 19,039"
Tu respuesta:
{"message":"Listo, dejé el saldo inicial de Santander en $19,039.","actions":[{"type":"edit_account","id":"<id real de Santander>","newInitialBalance":19039}]}

Usuario: "cambia el nombre de BBVA a Mi Banco"
Tu respuesta:
{"message":"Listo, BBVA ahora se llama Mi Banco.","actions":[{"type":"edit_account","id":"<id real de BBVA>","newName":"Mi Banco"}]}

Usuario: "el deposito de 7500 y el pago de seguro de 6800 son un reembolso, vincúlalos"
(asumiendo que en Movimientos recientes encuentras ingreso id="t-aa" de 7500 y gasto id="t-bb" de 6800)
Tu respuesta:
{"message":"Listo, los vinculé como reembolso. Solo cuenta como ingreso real la diferencia de $700.","actions":[{"type":"mark_passthrough","ids":["t-aa","t-bb"]}]}

Usuario: "el ingreso de los 7500 no es real, era un encargo"
(encuentras id="t-aa")
Tu respuesta:
{"message":"Marqué el ingreso de $7,500 como dinero de paso. No contará en tus estadísticas.","actions":[{"type":"mark_passthrough","ids":["t-aa"]}]}`;
}

/* persistencia — Firestore */
async function loadAll() {
  const u = auth.currentUser;
  if (!u) return { config: null, txs: [] };
  let config = null, txs = [];
  try {
    const snap = await getDoc(doc(db, "users", u.uid, "data", "config"));
    if (snap.exists()) config = snap.data().value;
  } catch (e) { console.error("loadAll config", e); }
  try {
    const snap = await getDoc(doc(db, "users", u.uid, "data", "txs"));
    if (snap.exists()) txs = snap.data().value || [];
  } catch (e) { console.error("loadAll txs", e); }
  return { config, txs };
}

async function persist(key, val) {
  const u = auth.currentUser;
  if (!u) return;
  const field = key === "cc:config" ? "config" : "txs";
  try {
    await setDoc(doc(db, "users", u.uid, "data", field), { value: val });
  } catch (e) { console.error("persist", e); }
}

/* llamada a Claude con imágenes (visión) */
async function callClaudeVision(system, userText, imagesB64) {
  const content = [
    ...imagesB64.map((b) => ({
      type: "image",
      source: { type: "base64", media_type: b.mediaType, data: b.data },
    })),
    { type: "text", text: userText },
  ];
  const body = { model: "claude-sonnet-4-6", max_tokens: 4000, system, messages: [{ role: "user", content }] };
  const isLocal = typeof window !== "undefined" && window.location.hostname === "localhost";
  let res;
  if (isLocal) {
    const key = import.meta.env.VITE_ANTHROPIC_KEY;
    res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01", "anthropic-dangerous-request-browser": "true" },
      body: JSON.stringify(body),
    });
  } else {
    res = await fetch("/api/claude", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  }
  if (!res.ok) throw new Error("vision " + res.status);
  const data = await res.json();
  return (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
}

/* leer File como base64 (sin el prefijo data:...) */
function fileToB64(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => {
      const result = String(r.result || "");
      const [meta, data] = result.split(",");
      const mediaType = (meta.match(/data:([^;]+)/) || [])[1] || "image/png";
      resolve({ mediaType, data });
    };
    r.onerror = () => reject(new Error("read failed"));
    r.readAsDataURL(file);
  });
}

/* llamada a Claude */
async function callClaude(system, messages) {
  const body = { model: "claude-sonnet-4-6", max_tokens: 1000, system, messages };
  const isLocal = typeof window !== "undefined" && window.location.hostname === "localhost";
  let res;
  if (isLocal) {
    const key = import.meta.env.VITE_ANTHROPIC_KEY;
    res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": key, "anthropic-version": "2023-06-01", "anthropic-dangerous-request-browser": "true" },
      body: JSON.stringify(body),
    });
  } else {
    res = await fetch("/api/claude", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  }
  if (!res.ok) throw new Error("api " + res.status);
  const data = await res.json();
  return (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n");
}
function parseJSON(text) {
  const clean = text.replace(/```json/gi, "").replace(/```/g, "").trim();
  const a = clean.indexOf("{"), b = clean.lastIndexOf("}");
  return JSON.parse(a >= 0 && b >= 0 ? clean.slice(a, b + 1) : clean);
}

/* Categorizador compartido: keywords locales -> Claude.
 * Devuelve { catId, sure }. catId=null si no encuentra nada seguro. */
async function autoCategorize(desc, type, accountId, config) {
  const nd = norm(desc);
  const cats = config.categories.filter((c) => c.type === type && c.accountId === accountId);
  // 1) claves locales de las categorías de ESTA cuenta
  for (const c of cats) {
    for (const kw of c.keywords || []) {
      if (nd.includes(norm(kw))) return { catId: c.id, sure: true };
    }
  }
  // 2) claves aprendidas en otras cuentas con mismo nombre/tipo
  for (const c of cats) {
    const twins = config.categories.filter(
      (x) => x.id !== c.id && x.type === c.type && norm(x.name).trim() === norm(c.name).trim()
    );
    for (const twin of twins) {
      for (const kw of twin.keywords || []) {
        if (nd.includes(norm(kw))) return { catId: c.id, sure: true };
      }
    }
  }
  // 3) Claude
  try {
    const sys =
      'Eres un clasificador de transacciones personales. Responde SOLO con un objeto JSON, sin markdown ni texto extra: {"categoryId":"<id>" o null,"confident":true|false}. Devuelve null y confident:false si no estás razonablemente seguro.';
    const user = `Tipo: ${type === "income" ? "ingreso" : "gasto"}. Descripción: "${desc}". Categorías disponibles: ${JSON.stringify(
      cats.map((c) => ({ id: c.id, name: c.name }))
    )}`;
    const raw = await callClaude(sys, [{ role: "user", content: user }]);
    const p = parseJSON(raw);
    const valid = cats.some((c) => c.id === p.categoryId);
    if (valid && p.confident) return { catId: p.categoryId, sure: true };
    if (valid) return { catId: p.categoryId, sure: false };
  } catch (e) { /* sin API */ }
  return { catId: null, sure: false };
}

/* categorías por defecto para una cuenta nueva */
function defaultCatsForAccount(accId) {
  return DEFAULT_CATS.map((t) => ({
    id: uid(), name: t.name, emoji: t.emoji, type: t.type, accountId: accId,
    keywords: SEED_KW[norm(t.name).trim()] ? [...SEED_KW[norm(t.name).trim()]] : [],
  }));
}

/* construir config completa desde el chatbot o el set default */
function buildConfig(raw) {
  const accountMode = raw.accountMode === "multiple" ? "multiple" : "single";
  let accounts = (raw.accounts || []).map((a) => ({
    id: uid(), name: a.name || "Cuenta", initialBalance: Number(a.initialBalance) || 0,
  }));
  if (!accounts.length) accounts = [{ id: uid(), name: "Principal", initialBalance: 0 }];

  let templates = (raw.categories || []).map((c) => ({
    name: c.name || "Categoría", emoji: c.emoji || "📦",
    type: c.type === "income" ? "income" : "expense",
  }));
  if (!templates.some((c) => c.type === "income")) templates.push({ name: "Otros ingresos", emoji: "💰", type: "income" });
  if (!templates.some((c) => c.type === "expense")) templates.push({ name: "Otros gastos", emoji: "📦", type: "expense" });

  // cada cuenta arranca con su propia copia de las categorías
  const categories = [];
  accounts.forEach((acc) => {
    templates.forEach((t) => {
      categories.push({
        id: uid(), name: t.name, emoji: t.emoji, type: t.type, accountId: acc.id,
        keywords: SEED_KW[norm(t.name).trim()] ? [...SEED_KW[norm(t.name).trim()]] : [],
      });
    });
  });
  return { setupComplete: true, accountMode, accounts, categories };
}

/* migración: convierte categorías globales (versión vieja) en categorías por cuenta */
function migrate(config, txs) {
  if (!config || !config.categories) return { config, txs };
  let outConfig = config;
  let outTxs = txs || [];

  // Migración 1: categorías globales -> categorías por cuenta
  if (config.categories.some((c) => !c.accountId)) {
    const oldCats = config.categories;
    const newCats = [];
    const map = {};
    config.accounts.forEach((acc) => {
      oldCats.forEach((oc) => {
        const nid = uid();
        newCats.push({
          id: nid, name: oc.name, emoji: oc.emoji || "📦",
          type: oc.type === "income" ? "income" : "expense",
          accountId: acc.id, keywords: [...(oc.keywords || [])],
        });
        map[`${acc.id}|${oc.id}`] = nid;
      });
    });
    outTxs = outTxs.map((t) => {
      const nid = map[`${t.accountId}|${t.categoryId}`];
      return nid ? { ...t, categoryId: nid } : t;
    });
    outConfig = { ...config, categories: newCats };
  }

  // Migración 2: linkedTxId (uno a uno) -> groupId (muchos a muchos)
  const hasLinked = outTxs.some((t) => t.linkedTxId);
  if (hasLinked) {
    const byId = new Map(outTxs.map((t) => [t.id, t]));
    const groupOf = new Map(); // id -> groupId
    function find(id) {
      while (groupOf.get(id) && groupOf.get(id) !== id) id = groupOf.get(id);
      return id;
    }
    outTxs.forEach((t) => { if (t.passThrough) groupOf.set(t.id, t.id); });
    outTxs.forEach((t) => {
      if (!t.passThrough || !t.linkedTxId) return;
      const other = byId.get(t.linkedTxId);
      if (!other || !other.passThrough) return;
      const r1 = find(t.id), r2 = find(other.id);
      if (r1 !== r2) groupOf.set(r2, r1);
    });
    outTxs = outTxs.map((t) => {
      if (!t.passThrough) return t;
      const root = groupOf.get(t.id) ? find(t.id) : null;
      const { linkedTxId, ...rest } = t;
      return root ? { ...rest, groupId: t.groupId || root } : rest;
    });
  }

  // Migración 3: quitar pass-through del MVP — los movimientos viejos vuelven a contar como reales
  const hadPassThrough = outTxs.some((t) => t.passThrough || t.linkedTxId || t.groupId);
  if (hadPassThrough) {
    outTxs = outTxs.map((t) => {
      const { passThrough, linkedTxId, groupId, ...rest } = t;
      return rest;
    });
  }

  return { config: outConfig, txs: outTxs };
}

/* ========================================================================= */
/* =========================================================================
   AUTH — pantallas de bienvenida, login, registro, recuperar contraseña
   ========================================================================= */

const AUTH_INK = "#1B2230";
const AUTH_INK_SOFT = "#6B7585";
const AUTH_GREEN = "#1A7A6E";
const AUTH_LINE = "#DDE2E9";
const AUTH_CORAL = "#B5453A";

const COUNTRIES = [
  { name: "Mexico", flag: "🇲🇽" },
  { name: "United States", flag: "🇺🇸" },
  { name: "Argentina", flag: "🇦🇷" },
  { name: "Bolivia", flag: "🇧🇴" },
  { name: "Brazil", flag: "🇧🇷" },
  { name: "Canada", flag: "🇨🇦" },
  { name: "Chile", flag: "🇨🇱" },
  { name: "Colombia", flag: "🇨🇴" },
  { name: "Costa Rica", flag: "🇨🇷" },
  { name: "Cuba", flag: "🇨🇺" },
  { name: "Ecuador", flag: "🇪🇨" },
  { name: "El Salvador", flag: "🇸🇻" },
  { name: "España", flag: "🇪🇸" },
  { name: "Guatemala", flag: "🇬🇹" },
  { name: "Honduras", flag: "🇭🇳" },
  { name: "Nicaragua", flag: "🇳🇮" },
  { name: "Panama", flag: "🇵🇦" },
  { name: "Paraguay", flag: "🇵🇾" },
  { name: "Peru", flag: "🇵🇪" },
  { name: "Puerto Rico", flag: "🇵🇷" },
  { name: "Dominican Republic", flag: "🇩🇴" },
  { name: "Uruguay", flag: "🇺🇾" },
  { name: "Venezuela", flag: "🇻🇪" },
];

function ZafiLogo() {
  return (
    <div style={{ display:"flex", flexDirection:"column", alignItems:"center", gap:8 }}>
      <span style={{ fontFamily:"'Fraunces',serif", fontWeight:400, fontSize:38,
        letterSpacing:"-.05em", color:AUTH_INK, fontFeatureSettings:'"ss01"', lineHeight:1 }}>
        zafi
      </span>
      <div style={{ fontFamily:"'Montserrat',sans-serif", fontSize:13, color:AUTH_INK_SOFT, letterSpacing:".01em", fontWeight:400 }}>
        Finanzas personales con IA
      </div>
    </div>
  );
}

function AuthScreen() {
  const [screen, setScreen] = useState("welcome");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [name, setName] = useState("");
  const [gender, setGender] = useState("");
  const [age, setAge] = useState("");
  const [country, setCountry] = useState("Mexico");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [ok, setOk] = useState("");

  function reset() { setEmail(""); setPassword(""); setConfirm(""); setName(""); setGender(""); setAge(""); setCountry("Mexico"); setErr(""); setOk(""); }
  function go(s) { reset(); setScreen(s); }

  function ferr(code) {
    return ({
      "auth/email-already-in-use": "Ya existe una cuenta con ese correo.",
      "auth/invalid-email": "El correo no es válido.",
      "auth/weak-password": "La contraseña debe tener al menos 6 caracteres.",
      "auth/user-not-found": "No encontramos una cuenta con ese correo.",
      "auth/wrong-password": "Contraseña incorrecta.",
      "auth/invalid-credential": "Correo o contraseña incorrectos.",
      "auth/too-many-requests": "Demasiados intentos. Espera un momento.",
      "auth/network-request-failed": "Sin conexión a internet.",
    })[code] || "Algo salió mal. Intenta de nuevo.";
  }

  async function doRegister() {
    if (!name.trim()) { setErr("Enter your name."); return; }
    if (!gender) { setErr("Select your gender."); return; }
    if (!age || Number(age) < 1 || Number(age) > 120) { setErr("Enter a valid age."); return; }
    if (!country.trim()) { setErr("Select your country."); return; }
    if (!email || !password || !confirm) { setErr("Fill in all fields."); return; }
    if (password !== confirm) { setErr("Passwords don't match."); return; }
    if (password.length < 6) { setErr("Password must be at least 6 characters."); return; }
    setBusy(true); setErr("");
    try {
      const cred = await createUserWithEmailAndPassword(auth, email, password);
      // guardar perfil dentro del config para que no vuelva a pedir ProfileSetup
      const avatarId = defaultAvatarForGender(gender);
      try {
        await setDoc(doc(db, "users", cred.user.uid, "data", "config"), {
          value: {
            userName: name.trim(), userGender: gender, userAge: Number(age),
            userCountry: country.trim(), ...(avatarId ? { avatarId } : {}),
          }
        }, { merge: true });
      } catch (e) { /* no bloquea el registro */ }
    }
    catch(e) { setErr(ferr(e.code)); }
    finally { setBusy(false); }
  }

  async function doLogin() {
    if (!email || !password) { setErr("Llena todos los campos."); return; }
    setBusy(true); setErr("");
    try { await signInWithEmailAndPassword(auth, email, password); }
    catch(e) { setErr(ferr(e.code)); }
    finally { setBusy(false); }
  }

  async function doForgot() {
    if (!email) { setErr("Escribe tu correo primero."); return; }
    setBusy(true); setErr(""); setOk("");
    try { await sendPasswordResetEmail(auth, email); setOk("Te enviamos un correo para restablecer tu contraseña."); }
    catch(e) { setErr(ferr(e.code)); }
    finally { setBusy(false); }
  }

  const FONT = "'Montserrat', sans-serif";
  const inp = {
    width:"100%", padding:"14px 16px", borderRadius:16,
    border:"1px solid rgba(255,255,255,.5)", fontSize:15, fontFamily:FONT, fontWeight:400,
    background:"rgba(255,255,255,.55)", color:AUTH_INK, outline:"none",
    backdropFilter:"blur(10px)", WebkitBackdropFilter:"blur(10px)",
  };
  const btnP = {
    width:"100%", padding:15, borderRadius:99, border:"none",
    background:AUTH_INK, color:"#fff", fontSize:15, fontWeight:600,
    fontFamily:FONT, cursor:busy?"not-allowed":"pointer", opacity:busy?.6:1,
    letterSpacing:"-.01em", transition:"transform .1s ease",
  };
  const btnS = {
    width:"100%", padding:14, borderRadius:99,
    border:"1px solid rgba(255,255,255,.5)", background:"rgba(255,255,255,.45)",
    backdropFilter:"blur(10px)", WebkitBackdropFilter:"blur(10px)",
    color:AUTH_INK, fontSize:15, fontWeight:500,
    fontFamily:FONT, cursor:"pointer", letterSpacing:"-.01em",
  };
  const lnk = {
    background:"none", border:"none", color:AUTH_GREEN,
    fontSize:14, fontFamily:FONT, cursor:"pointer",
    fontWeight:600, padding:0,
  };
  const lbl = {
    fontFamily:FONT, fontSize:12, fontWeight:600, color:AUTH_INK_SOFT,
    letterSpacing:".02em", marginBottom:6, display:"block",
  };
  const wrap = {
    minHeight:"100vh", display:"flex", flexDirection:"column",
    alignItems:"center", justifyContent:"center", padding:"40px 24px",
    background:"#DCE1E8", position:"relative", overflow:"hidden",
  };
  const blob = {
    position:"absolute", width:280, height:280, borderRadius:"50%",
    background:"radial-gradient(circle, rgba(96,165,250,.25) 0%, rgba(96,165,250,.05) 60%, transparent 80%)",
    filter:"blur(40px)", top:"15%", right:"-10%", pointerEvents:"none",
  };
  const blob2 = {
    position:"absolute", width:200, height:200, borderRadius:"50%",
    background:"radial-gradient(circle, rgba(167,139,250,.18) 0%, rgba(167,139,250,.04) 60%, transparent 80%)",
    filter:"blur(30px)", bottom:"20%", left:"-5%", pointerEvents:"none",
  };
  const box = {
    width:"100%", maxWidth:360, display:"flex", flexDirection:"column", gap:24,
  };

  if (screen === "welcome") return (
    <div style={{ minHeight:"100vh", position:"relative", overflow:"hidden", background:"#DCE1E8",
      display:"flex", flexDirection:"column", justifyContent:"flex-end" }}>
      {/* video de fondo */}
      <video
        ref={(el) => {
          if (el) {
            el.muted = true;
            const p = el.play();
            if (p && p.catch) p.catch(() => {});
          }
        }}
        autoPlay muted loop playsInline preload="auto"
        style={{ position:"absolute", inset:0, width:"100%", height:"100%",
          objectFit:"cover", zIndex:0, background:"#DCE1E8" }}
      >
        <source src="/zafi-intro.mp4" type="video/mp4" />
      </video>
      {/* degradado para legibilidad abajo */}
      <div style={{ position:"absolute", inset:0, zIndex:1, pointerEvents:"none",
        background:"linear-gradient(to bottom, rgba(220,225,232,0) 35%, rgba(220,225,232,.55) 70%, rgba(220,225,232,.85) 100%)" }} />

      {/* contenido abajo */}
      <div style={{ position:"relative", zIndex:2, padding:"0 26px calc(34px + env(safe-area-inset-bottom))",
        display:"flex", flexDirection:"column", gap:22 }}>
        <div>
          <div style={{ fontFamily:"'Fraunces',serif", fontWeight:400, fontSize:56,
            letterSpacing:"-.05em", color:AUTH_INK, lineHeight:1, fontFeatureSettings:'"ss01"' }}>zafi</div>
          <div style={{ fontFamily:"'Montserrat',sans-serif", fontWeight:400, fontSize:16,
            color:AUTH_INK_SOFT, marginTop:6, letterSpacing:"-.01em" }}>
            finanzas personales con <span style={{ fontFamily:"'Fraunces',serif", fontStyle:"italic" }}>IA</span>
          </div>
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          <button
            style={{ width:"100%", padding:16, borderRadius:99, border:"1px solid rgba(255,255,255,.6)",
              background:"rgba(255,255,255,.65)", backdropFilter:"blur(12px)", WebkitBackdropFilter:"blur(12px)",
              color:AUTH_INK, fontSize:15, fontWeight:600, fontFamily:"'Montserrat', sans-serif", cursor:"pointer",
              letterSpacing:"-.01em", boxShadow:"0 6px 24px rgba(30,40,60,.12)" }}
            onClick={() => go("method")}>Crear cuenta</button>
          <button
            style={{ width:"100%", padding:16, borderRadius:99, border:"1px solid rgba(255,255,255,.25)",
              background:"rgba(27,34,48,.55)", backdropFilter:"blur(12px)", WebkitBackdropFilter:"blur(12px)",
              color:"#fff", fontSize:15, fontWeight:500, fontFamily:"'Montserrat', sans-serif", cursor:"pointer",
              letterSpacing:"-.01em", boxShadow:"0 6px 24px rgba(30,40,60,.18)" }}
            onClick={() => go("login")}>Iniciar sesión</button>
        </div>
        <div style={{ fontSize:11.5, color:AUTH_INK_SOFT, textAlign:"center", lineHeight:1.5, opacity:.75, fontFamily:"'Montserrat', sans-serif" }}>
          Al crear una cuenta aceptas nuestros términos de uso y política de privacidad.
        </div>
      </div>
    </div>
  );

  // ===== Shared step page wrapper =====
  const stepWrap = {
    minHeight:"100vh", display:"flex", flexDirection:"column",
    background:"#EAEEF4", position:"relative", fontFamily:FONT,
  };
  const backBtn = {
    position:"absolute", top:20, left:20, zIndex:2, width:44, height:44,
    borderRadius:"50%", border:"none", background:"#fff",
    boxShadow:"0 2px 10px rgba(30,40,60,.1)", cursor:"pointer",
    fontSize:20, display:"flex", alignItems:"center", justifyContent:"center",
    color:AUTH_INK, fontFamily:FONT,
  };
  const stepBottom = {
    padding:"0 26px calc(28px + env(safe-area-inset-bottom))",
    display:"flex", flexDirection:"column", gap:12,
  };

  // ===== PASO 1: Método (Continue with email / Google / Apple) =====
  if (screen === "method") return (
    <div style={{ ...stepWrap, justifyContent:"flex-end" }}>
      <button style={backBtn} onClick={() => go("welcome")}>‹</button>
      <div style={{ flex:1 }} />
      <div style={{ padding:"0 26px 32px" }}>
        <div style={{ fontFamily:"'Fraunces',serif", fontWeight:400, fontSize:64,
          letterSpacing:"-.05em", color:AUTH_INK, lineHeight:1, fontFeatureSettings:'"ss01"' }}>zafi</div>
      </div>
      <div style={stepBottom}>
        <button style={{ ...btnP, borderRadius:99 }} onClick={() => { setErr(""); setScreen("country"); }}>
          Continue with your email
        </button>
        <button style={{ width:"100%", padding:15, borderRadius:99, border:"none",
          background:"#4285F4", color:"#fff", fontSize:15, fontWeight:600,
          fontFamily:FONT, cursor:"pointer", letterSpacing:"-.01em",
          display:"flex", alignItems:"center", justifyContent:"center", gap:10 }}
          onClick={async () => {
            setBusy(true); setErr("");
            try {
              const gp = new GoogleAuthProvider();
              await signInWithPopup(auth, gp);
            }
            catch(e) { setErr(ferr(e.code)); }
            setBusy(false);
          }}
          disabled={busy}>
          Continue with Google
          <span style={{ fontSize:18, fontWeight:700 }}>G</span>
        </button>
        <button style={{ width:"100%", padding:15, borderRadius:99, border:"none",
          background:"#000", color:"#fff", fontSize:15, fontWeight:600,
          fontFamily:FONT, cursor:"pointer", letterSpacing:"-.01em",
          display:"flex", alignItems:"center", justifyContent:"center", gap:10 }}
          onClick={async () => {
            setBusy(true); setErr("");
            try {
              const ap = new OAuthProvider("apple.com");
              await signInWithPopup(auth, ap);
            }
            catch(e) { setErr(ferr(e.code)); }
            setBusy(false);
          }}
          disabled={busy}>
          Continue with Apple
          <span style={{ fontSize:18 }}></span>
        </button>
        {err && <div style={{ fontSize:13, color:AUTH_CORAL, fontWeight:500, textAlign:"center" }}>{err}</div>}
      </div>
    </div>
  );

  // ===== PASO 2: País =====
  const selCountry = COUNTRIES.find(c => c.name === country) || COUNTRIES[0];
  if (screen === "country") return (
    <div style={{ ...stepWrap }}>
      <button style={backBtn} onClick={() => { setErr(""); setScreen("method"); }}>‹</button>
      <div style={{ flex:1, padding:"90px 26px 0" }}>
        <div style={{ fontSize:28, fontWeight:700, color:AUTH_INK, letterSpacing:"-.02em", lineHeight:1.2 }}>
          Country of Residence
        </div>
        <div style={{ fontSize:14, color:AUTH_INK_SOFT, marginTop:8, lineHeight:1.6 }}>
          We need to collect a few personal details to verify you're elegible for an account
        </div>
        <div style={{ marginTop:32 }}>
          <div style={{ fontSize:13, color:AUTH_INK_SOFT, marginBottom:8 }}>Country</div>
          <div style={{ position:"relative" }}>
            <select
              value={country}
              onChange={e => setCountry(e.target.value)}
              style={{ width:"100%", padding:"12px 0", fontSize:16, fontWeight:600,
                fontFamily:FONT, color:"transparent", background:"transparent",
                border:"none", borderBottom:"1px solid rgba(27,34,48,.15)",
                outline:"none", appearance:"none", cursor:"pointer",
                position:"relative", zIndex:1 }}>
              {COUNTRIES.map(c => (
                <option key={c.name} value={c.name}>{c.flag}  {c.name}</option>
              ))}
            </select>
            <div style={{ position:"absolute", top:12, left:0, fontSize:16, fontWeight:600, color:AUTH_INK, pointerEvents:"none" }}>
              {selCountry.flag} {selCountry.name}
            </div>
          </div>
        </div>
      </div>
      <div style={stepBottom}>
        <div style={{ fontSize:13, color:AUTH_INK_SOFT, lineHeight:1.6, marginBottom:4 }}>
          By tapping Get started, you agree to Zafi's <span style={{ textDecoration:"underline", cursor:"pointer" }}>Terms and Conditions</span> and <span style={{ textDecoration:"underline", cursor:"pointer" }}>Privacy Policy</span>, and represent that you are acting on your own behalf.
        </div>
        <button style={{ ...btnP, borderRadius:99 }} onClick={() => { setErr(""); setScreen("profile"); }}>
          Get started
        </button>
      </div>
    </div>
  );

  // ===== PASO 3: Datos personales (Nombre, Edad, Género) =====
  if (screen === "profile") return (
    <div style={{ ...stepWrap }}>
      <button style={backBtn} onClick={() => { setErr(""); setScreen("country"); }}>‹</button>
      <div style={{ flex:1, padding:"90px 26px 0" }}>
        <div style={{ fontSize:28, fontWeight:700, color:AUTH_INK, letterSpacing:"-.02em", lineHeight:1.2 }}>
          About you
        </div>
        <div style={{ fontSize:14, color:AUTH_INK_SOFT, marginTop:8, lineHeight:1.6 }}>
          Tell us a bit about yourself so we can personalize your experience.
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:18, marginTop:28 }}>
          <div>
            <label style={lbl}>Name</label>
            <input style={{ ...inp, background:"rgba(255,255,255,.7)" }} type="text" placeholder="Your name" value={name} onChange={e=>setName(e.target.value)} />
          </div>
          <div>
            <label style={lbl}>Age</label>
            <input style={{ ...inp, background:"rgba(255,255,255,.7)", width:120 }} type="text" inputMode="numeric" placeholder="00" value={age}
              onChange={e=>setAge(e.target.value.replace(/[^0-9]/g,"").slice(0,3))} />
          </div>
          <div>
            <label style={lbl}>Gender</label>
            <div style={{ display:"flex", gap:9 }}>
              {[["male","Male"],["female","Female"],["other","Other"]].map(([k,l])=>(
                <button key={k} type="button" onClick={()=>setGender(k)}
                  style={{ flex:1, padding:"13px 8px", borderRadius:14, cursor:"pointer", fontFamily:FONT,
                    fontSize:14, fontWeight:gender===k?600:400,
                    background: gender===k ? AUTH_INK : "rgba(255,255,255,.7)",
                    color: gender===k ? "#fff" : AUTH_INK,
                    border:`1px solid ${gender===k ? AUTH_INK : "rgba(0,0,0,.08)"}`,
                    transition:"background .15s, color .15s" }}>
                  {l}
                </button>
              ))}
            </div>
          </div>
        </div>
        {err && <div style={{ fontSize:13, color:AUTH_CORAL, fontWeight:500, marginTop:12 }}>{err}</div>}
      </div>
      <div style={stepBottom}>
        <button style={{ ...btnP, borderRadius:99 }}
          onClick={() => {
            if (!name.trim()) { setErr("Enter your name."); return; }
            if (!age || Number(age) < 1 || Number(age) > 120) { setErr("Enter a valid age."); return; }
            if (!gender) { setErr("Select your gender."); return; }
            setErr(""); setScreen("credentials");
          }}>
          Continue
        </button>
      </div>
    </div>
  );

  // ===== PASO 4: Credenciales (Email + Contraseña) =====
  if (screen === "credentials") return (
    <div style={{ ...stepWrap }}>
      <button style={backBtn} onClick={() => { setErr(""); setScreen("profile"); }}>‹</button>
      <div style={{ flex:1, padding:"90px 26px 0" }}>
        <div style={{ fontSize:28, fontWeight:700, color:AUTH_INK, letterSpacing:"-.02em", lineHeight:1.2 }}>
          Create your account
        </div>
        <div style={{ fontSize:14, color:AUTH_INK_SOFT, marginTop:8, lineHeight:1.6 }}>
          You'll use this email and password to sign in.
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:16, marginTop:28 }}>
          <div>
            <label style={lbl}>Email</label>
            <input style={{ ...inp, background:"rgba(255,255,255,.7)" }} type="email" placeholder="you@example.com" value={email} onChange={e=>setEmail(e.target.value)} />
          </div>
          <div>
            <label style={lbl}>Password</label>
            <input style={{ ...inp, background:"rgba(255,255,255,.7)" }} type="password" placeholder="At least 6 characters" value={password} onChange={e=>setPassword(e.target.value)} />
          </div>
          <div>
            <label style={lbl}>Confirm password</label>
            <input style={{ ...inp, background:"rgba(255,255,255,.7)" }} type="password" placeholder="Repeat your password" value={confirm} onChange={e=>setConfirm(e.target.value)} />
          </div>
        </div>
        {err && <div style={{ fontSize:13, color:AUTH_CORAL, fontWeight:500, marginTop:12 }}>{err}</div>}
      </div>
      <div style={stepBottom}>
        <button style={{ ...btnP, borderRadius:99 }} onClick={doRegister} disabled={busy}>
          {busy ? "Creating account…" : "Create account"}
        </button>
      </div>
    </div>
  );

  if (screen === "login") return (
    <div style={wrap}>
      <div style={blob} />
      <div style={blob2} />
      <div style={{ ...box, position:"relative", zIndex:1 }}>
        <ZafiLogo />
        <div>
          <div style={{ fontFamily:FONT, fontSize:24, fontWeight:700, color:AUTH_INK, letterSpacing:"-.02em", marginBottom:4 }}>Bienvenido de vuelta</div>
          <div style={{ fontFamily:FONT, fontSize:14, fontWeight:400, color:AUTH_INK_SOFT }}>Inicia sesión para continuar.</div>
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
          <div>
            <label style={lbl}>Correo electrónico</label>
            <input style={inp} type="email" placeholder="tucorreo@ejemplo.com" value={email} onChange={e=>setEmail(e.target.value)} />
          </div>
          <div>
            <label style={lbl}>Contraseña</label>
            <input style={inp} type="password" placeholder="Tu contraseña" value={password} onChange={e=>setPassword(e.target.value)} />
          </div>
        </div>
        {err && <div style={{ fontFamily:FONT, fontSize:13.5, color:AUTH_CORAL, fontWeight:500 }}>{err}</div>}
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          <button style={btnP} onClick={doLogin} disabled={busy}>{busy?"Entrando…":"Iniciar sesión"}</button>
          <button style={btnS} onClick={() => go("welcome")}>← Regresar</button>
        </div>
        <div style={{ textAlign:"center", fontSize:14, color:AUTH_INK_SOFT, display:"flex", flexDirection:"column", gap:8 }}>
          <button style={lnk} onClick={() => go("forgot")}>¿Olvidaste tu contraseña?</button>
          <span>¿No tienes cuenta?{" "}<button style={lnk} onClick={() => go("method")}>Crear cuenta</button></span>
        </div>
      </div>
    </div>
  );

  if (screen === "forgot") return (
    <div style={wrap}>
      <div style={blob} />
      <div style={blob2} />
      <div style={{ ...box, position:"relative", zIndex:1 }}>
        <ZafiLogo />
        <div>
          <div style={{ fontFamily:FONT, fontSize:24, fontWeight:700, color:AUTH_INK, letterSpacing:"-.02em", marginBottom:4 }}>Restablecer contraseña</div>
          <div style={{ fontFamily:FONT, fontSize:14, fontWeight:400, color:AUTH_INK_SOFT, lineHeight:1.5 }}>Escribe tu correo y te mandamos un enlace para crear una nueva contraseña.</div>
        </div>
        <div>
          <label style={lbl}>Correo electrónico</label>
          <input style={inp} type="email" placeholder="tucorreo@ejemplo.com" value={email} onChange={e=>setEmail(e.target.value)} />
        </div>
        {err && <div style={{ fontFamily:FONT, fontSize:13.5, color:AUTH_CORAL, fontWeight:500 }}>{err}</div>}
        {ok && <div style={{ fontFamily:FONT, fontSize:13.5, color:AUTH_GREEN, fontWeight:500 }}>{ok}</div>}
        <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
          <button style={btnP} onClick={doForgot} disabled={busy}>{busy?"Enviando…":"Enviar correo"}</button>
          <button style={btnS} onClick={() => go("login")}>← Regresar</button>
        </div>
      </div>
    </div>
  );
}

/* =========================================================================
   APP — componente principal
   ========================================================================= */
/* ===================== PROFILE SETUP (post-auth for Google/Apple) ======= */
function ProfileSetup({ user, config, saveConfig, onDone }) {
  const FONT = "'Montserrat', sans-serif";
  const [name, setName] = useState("");
  const [age, setAge] = useState("");
  const [gender, setGender] = useState("");
  const [country, setCountry] = useState("Mexico");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const selCountry = COUNTRIES.find(c => c.name === country) || COUNTRIES[0];

  const save = async () => {
    if (!name.trim()) { setErr("Enter your name."); return; }
    if (!age || Number(age) < 1) { setErr("Enter a valid age."); return; }
    if (!gender) { setErr("Select your gender."); return; }
    setBusy(true); setErr("");
    try {
      const avatarId = defaultAvatarForGender(gender);
      const profileData = {
        userName: name.trim(), userGender: gender, userAge: Number(age),
        userCountry: country, ...(avatarId ? { avatarId } : {}),
      };
      // merge into existing config or create minimal config
      const updated = { ...(config || {}), ...profileData };
      saveConfig(updated);
      onDone();
    } catch (e) { setErr("Something went wrong. Try again."); }
    setBusy(false);
  };

  const stepWrap = {
    minHeight:"100vh", display:"flex", flexDirection:"column",
    background:"#EAEEF4", fontFamily:FONT,
  };
  const lbl = { fontFamily:FONT, fontSize:12, fontWeight:600, color:"#6B7585", letterSpacing:".02em", marginBottom:6, display:"block" };
  const inp = { width:"100%", padding:"14px 16px", borderRadius:16,
    border:"1px solid rgba(0,0,0,.06)", fontSize:15, fontFamily:FONT, fontWeight:400,
    background:"rgba(255,255,255,.7)", color:"#1B2230", outline:"none" };

  return (
    <div style={stepWrap}>
      <div style={{ flex:1, padding:"100px 26px 0" }}>
        <div style={{ fontSize:28, fontWeight:700, color:"#1B2230", letterSpacing:"-.02em", lineHeight:1.2 }}>
          Welcome to Zafi
        </div>
        <div style={{ fontSize:14, color:"#6B7585", marginTop:8, lineHeight:1.6 }}>
          Tell us a bit about yourself so we can personalize your experience.
        </div>
        <div style={{ display:"flex", flexDirection:"column", gap:18, marginTop:28 }}>
          <div>
            <label style={lbl}>Your name</label>
            <input style={inp} type="text" placeholder="How should we call you?" value={name} onChange={e=>setName(e.target.value)} />
          </div>
          <div>
            <label style={lbl}>Age</label>
            <input style={{ ...inp, width:120 }} type="text" inputMode="numeric" placeholder="00" value={age}
              onChange={e=>setAge(e.target.value.replace(/[^0-9]/g,"").slice(0,3))} />
          </div>
          <div>
            <label style={lbl}>Gender</label>
            <div style={{ display:"flex", gap:9 }}>
              {[["male","Male"],["female","Female"],["other","Other"]].map(([k,l])=>(
                <button key={k} type="button" onClick={()=>setGender(k)}
                  style={{ flex:1, padding:"13px 8px", borderRadius:14, cursor:"pointer", fontFamily:FONT,
                    fontSize:14, fontWeight:gender===k?600:400,
                    background: gender===k ? "#1B2230" : "rgba(255,255,255,.7)",
                    color: gender===k ? "#fff" : "#1B2230",
                    border:`1px solid ${gender===k ? "#1B2230" : "rgba(0,0,0,.08)"}`,
                    transition:"background .15s, color .15s" }}>
                  {l}
                </button>
              ))}
            </div>
          </div>
          <div>
            <label style={lbl}>Country</label>
            <div style={{ position:"relative" }}>
              <select value={country} onChange={e => setCountry(e.target.value)}
                style={{ width:"100%", padding:"12px 0", fontSize:16, fontWeight:600,
                  fontFamily:FONT, color:"transparent", background:"transparent",
                  border:"none", borderBottom:"1px solid rgba(27,34,48,.15)",
                  outline:"none", appearance:"none", cursor:"pointer", position:"relative", zIndex:1 }}>
                {COUNTRIES.map(c => <option key={c.name} value={c.name}>{c.flag}  {c.name}</option>)}
              </select>
              <div style={{ position:"absolute", top:12, left:0, fontSize:16, fontWeight:600, color:"#1B2230", pointerEvents:"none" }}>
                {selCountry.flag} {selCountry.name}
              </div>
            </div>
          </div>
        </div>
        {err && <div style={{ fontSize:13, color:"#B5453A", fontWeight:500, marginTop:12 }}>{err}</div>}
      </div>
      <div style={{ padding:"0 26px calc(28px + env(safe-area-inset-bottom))" }}>
        <button style={{ width:"100%", padding:15, borderRadius:99, border:"none",
          background:"#1B2230", color:"#fff", fontSize:15, fontWeight:600,
          fontFamily:FONT, cursor:busy?"not-allowed":"pointer", opacity:busy?.6:1 }}
          onClick={save} disabled={busy}>
          {busy ? "Saving…" : "Continue"}
        </button>
      </div>
    </div>
  );
}

/* ===================== SPLASH SCREEN ==================================== */
function SplashScreen({ onDone }) {
  useEffect(() => {
    const t = setTimeout(onDone, 2100);
    return () => clearTimeout(t);
  }, [onDone]);

  return (
    <div className="cc-splash">
      <style>{STYLE}</style>
      <div className="cc-splash-orb" />
      <div className="cc-splash-word">zafi</div>
      <div className="cc-splash-tag">finanzas con IA</div>
    </div>
  );
}

export default function App() {
  const [splashDone, setSplashDone] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const [config, setConfig] = useState(null);
  const [txs, setTxs] = useState([]);
  const [toast, setToast] = useState(null);
  const [user, setUser] = useState(undefined); // undefined=cargando, null=no logueado
  const [profileDone, setProfileDone] = useState(false); // did user complete profile?

  useEffect(() => {
    if (typeof document !== "undefined") document.title = "Zafi · Finanzas personales con IA";
  }, []);

  // Escuchar sesión de Firebase
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => setUser(u));
    return () => unsub();
  }, []);

  // Cargar datos locales — siempre se define, solo corre cuando hay usuario
  useEffect(() => {
    if (!user) { setProfileDone(false); setLoaded(false); return; }
    let cancelled = false;
    (async () => {
      try {
        const { config: c, txs: t } = await loadAll();
        if (cancelled) return;
        if (c) {
          const m = migrate(c, t || []);
          const r = processRecurring(m.config, m.txs);
          setConfig(r.config);
          setTxs(r.txs);
          const configChanged = m.config !== c || r.generated > 0;
          const txsChanged = m.txs !== t || r.generated > 0;
          if (configChanged) persist("cc:config", r.config);
          if (txsChanged) persist("cc:txs", r.txs);
          if (r.generated > 0) {
            setTimeout(() => {
              setToast(`Se aplicaron ${r.generated} movimiento${r.generated === 1 ? "" : "s"} de recurrencias`);
              setTimeout(() => setToast(null), 2800);
            }, 600);
          }
        } else if (t) {
          setTxs(t);
        }
        // Check if user has completed profile (stored in config)
        // Also accept existing users who have setupComplete but no userName yet
        const hasProfile = !!(c && (c.userName || c.setupComplete));
        if (!cancelled) setProfileDone(hasProfile);
      } catch (e) {
        console.error("load error", e);
        if (!cancelled) setProfileDone(false);
      }
      if (!cancelled) setLoaded(true);
    })();
    return () => { cancelled = true; };
  }, [user]); // re-corre cuando cambia el usuario

  // Splash de bienvenida — siempre primero al abrir la app
  if (!splashDone) return <SplashScreen onDone={() => setSplashDone(true)} />;

  // Pantalla de carga mientras Firebase verifica sesión
  if (user === undefined) return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center",
      justifyContent:"center", background:"#DCE1E8" }}>
      <div style={{ display:"flex", flexDirection:"column", alignItems:"center", gap:12 }}>
        <span style={{ fontFamily:"'Fraunces',serif", fontWeight:400, fontSize:26,
          letterSpacing:"-.05em", color:"#1B2230" }}>zafi</span>
        <div style={{ fontSize:13, color:"#A4ACBA" }}>Cargando…</div>
      </div>
    </div>
  );

  // Sin sesión → pantalla de Auth
  if (!user) return <AuthScreen />;

  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2400);
  };
  const saveConfig = (c) => { setConfig(c); persist("cc:config", c); };
  const saveTxs = (t) => { setTxs(t); persist("cc:txs", t); };
  const resetAll = async () => {
    const u = auth.currentUser;
    if (u) {
      try { await deleteDoc(doc(db, "users", u.uid, "data", "config")); } catch (e) {}
      try { await deleteDoc(doc(db, "users", u.uid, "data", "txs")); } catch (e) {}
    }
    setConfig(null);
    setTxs([]);
    showToast("Listo, empezamos desde cero");
  };

  if (!loaded)
    return (
      <div className="cc-root" style={{ display: "flex", alignItems: "center", justifyContent: "center", height: "100vh" }}>
        <style>{STYLE}</style>
        <div className="cc-video-bg">
          <video autoPlay muted loop playsInline preload="auto"
            ref={(el) => { if (el) { el.muted = true; const p = el.play(); if (p && p.catch) p.catch(() => {}); } }}>
            <source src="/zafi-bg.mp4" type="video/mp4" />
          </video>
        </div>
        <div className="cc-dots"><span /><span /><span /></div>
      </div>
    );

  // Si el usuario no tiene perfil (Google/Apple sin nombre), pedir datos
  if (!profileDone)
    return <ProfileSetup user={user} config={config} saveConfig={saveConfig} onDone={() => setProfileDone(true)} />;

  return (
    <div className="cc-root">
      <style>{STYLE}</style>
      <div className="cc-video-bg">
        <video autoPlay muted loop playsInline preload="auto"
          ref={(el) => { if (el) { el.muted = true; const p = el.play(); if (p && p.catch) p.catch(() => {}); } }}>
          <source src="/zafi-bg.mp4" type="video/mp4" />
        </video>
      </div>
      <div className="cc-bg-wave" />
      {!config?.setupComplete ? (
        <Onboarding onDone={(built) => saveConfig({ ...config, ...built })} />
      ) : (
        <Main config={config} txs={txs} saveConfig={saveConfig} saveTxs={saveTxs} showToast={showToast} resetAll={resetAll} />
      )}
      {toast && <div className="cc-toast">{toast}</div>}
    </div>
  );
}

/* ============================= ONBOARDING ================================ */
function Onboarding({ onDone }) {
  const GREETING =
    "¡Hola! Bienvenido a Zafi 👋\n\nSoy tu asistente. Te ayudaré a configurar la app en menos de un minuto, platicando aquí mismo.\n\n¿Quieres que use un set de categorías recomendado, o prefieres armarlas tú?";
  const [msgs, setMsgs] = useState([{ role: "bot", text: GREETING }]);
  const [suggs, setSuggs] = useState(["Usa las recomendadas", "Quiero armarlas yo"]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const [apiOk, setApiOk] = useState(true);
  const history = useRef([]); // [{role, content}]
  const scroller = useRef(null);

  useEffect(() => {
    if (scroller.current) scroller.current.scrollTop = scroller.current.scrollHeight;
  }, [msgs, busy]);

  const SYSTEM = `Eres el asistente de configuración de "Zafi", una app de finanzas personales con IA. El usuario habla español de México. Conversa de forma cálida, breve y natural.

Debes averiguar 2 cosas:
1) CATEGORÍAS para clasificar ingresos y gastos. Set recomendado: Sueldo, Negocio, Freelance, Otros ingresos (ingresos); Súper/Despensa, Restaurantes, Transporte/Gasolina, Casa/Renta, Servicios, Salud, Entretenimiento, Suscripciones, Otros gastos (gastos).
2) CUENTAS: si maneja una sola cuenta para todo, o varias (efectivo, banco, tarjeta, etc.).

REGLAS DE RESPUESTA:
- Responde SIEMPRE con UN SOLO objeto JSON válido, sin markdown ni texto fuera del JSON.
- Formato: {"message":"texto para el usuario","suggestions":["opción corta",...],"done":false,"config":null}
- "message": 2-4 frases máximo, UNA pregunta a la vez.
- "suggestions": hasta 4 respuestas rápidas cortas que el usuario puede tocar. Omite o usa [] si no aplica.
- Cuando ya tengas categorías y modo de cuentas definidos, pon "done":true y llena "config":
  {"accountMode":"single"|"multiple","accounts":[{"name":"...","initialBalance":<número: dinero que ya tiene en esa cuenta>}],"categories":[{"name":"...","emoji":"emoji","type":"income"|"expense"}]}
- En modo "single" usa una cuenta llamada "Principal" salvo que el usuario diga otro nombre.
- Si el usuario maneja varias cuentas, pregúntale el saldo inicial (dinero que tiene ahorita) de cada una y ponlo en "initialBalance". Si no lo sabe, usa 0.
- Cada categoría necesita un emoji apropiado.
- Si el usuario pide usar lo recomendado, ve directo a done:true con el set recomendado y pregunta el modo de cuentas si aún no lo sabes.`;

  async function send(text) {
    const userText = (text ?? input).trim();
    if (!userText || busy) return;
    setMsgs((m) => [...m, { role: "me", text: userText }]);
    setInput("");
    setSuggs([]);
    setBusy(true);
    history.current.push({ role: "user", content: userText });
    try {
      const raw = await callClaude(SYSTEM, history.current);
      const parsed = parseJSON(raw);
      history.current.push({ role: "assistant", content: parsed.message || raw });
      setMsgs((m) => [...m, { role: "bot", text: parsed.message || "Listo." }]);
      setSuggs(Array.isArray(parsed.suggestions) ? parsed.suggestions.slice(0, 4) : []);
      if (parsed.done && parsed.config) {
        setTimeout(() => onDone(buildConfig(parsed.config)), 700);
      }
    } catch (e) {
      setApiOk(false);
      setMsgs((m) => [
        ...m,
        { role: "bot", text: "Tuve un problema para conectar con la IA. No hay bronca: puedes usar la configuración recomendada con el botón de abajo. 👇" },
      ]);
    }
    setBusy(false);
  }

  return (
    <div>
      <div className="cc-top">
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-start", gap: 2 }}>
          <div className="cc-logo"><span className="cc-logo-dot" />zafi</div>
          <div style={{ fontSize: 11, color: "var(--ink-soft)", marginLeft: 19, letterSpacing: "-.005em" }}>
            Finanzas personales con IA
          </div>
        </div>
      </div>
      <div className="cc-wrap" style={{ paddingBottom: 40 }}>
        <div
          ref={scroller}
          style={{ display: "flex", flexDirection: "column", gap: 10, padding: "8px 2px 14px", maxHeight: "58vh", overflowY: "auto" }}
        >
          {msgs.map((m, i) => (
            <div key={i} className={`cc-bubble ${m.role} cc-fade`} style={{ whiteSpace: "pre-wrap" }}>
              {m.text}
            </div>
          ))}
          {busy && (
            <div className="cc-bubble bot"><span className="cc-dots"><span /><span /><span /></span></div>
          )}
        </div>

        {suggs.length > 0 && !busy && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 12 }}>
            {suggs.map((s, i) => (
              <button key={i} className="cc-chip" onClick={() => send(s)}>{s}</button>
            ))}
          </div>
        )}

        <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
          <input
            className="cc-input"
            placeholder="Escribe tu respuesta…"
            value={input}
            disabled={busy}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && send()}
          />
          <button className="cc-btn cc-btn-primary" disabled={busy || !input.trim()} onClick={() => send()}>
            Enviar
          </button>
        </div>

        <div style={{ textAlign: "center" }}>
          <button
            className="cc-btn"
            style={{ fontSize: 13 }}
            onClick={() => onDone(buildConfig({ accountMode: "single", accounts: [{ name: "Principal" }], categories: DEFAULT_CATS }))}
          >
            {apiOk ? "Saltar — usar configuración recomendada" : "Usar configuración recomendada ahora"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* =============================== MAIN ==================================== */
function Main({ config, txs, saveConfig, saveTxs, showToast, resetAll }) {
  setAppLang(config.language || "es");
  const [tab, setTab] = useState("inicio");
  const [adding, setAdding] = useState(false);
  const [editingTx, setEditingTx] = useState(null);
  const [chatOpen, setChatOpen] = useState(false);
  const [chatVoice, setChatVoice] = useState(false);
  const [importOpen, setImportOpen] = useState(false);
  const [addMenuOpen, setAddMenuOpen] = useState(false);
  const [accountsOpen, setAccountsOpen] = useState(false);
  const [excelOpen, setExcelOpen] = useState(false);
  const [rangeOpen, setRangeOpen] = useState(false);
  const [recurringOpen, setRecurringOpen] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);

  const dateRange = config.dateRange || DEFAULT_RANGE;
  const setDateRange = (newRange) => {
    saveConfig({ ...config, dateRange: newRange });
  };

  const balance = grandTotal(config, txs);

  // === Motor de recurrentes: al cargar, genera movimientos pendientes ===
  const recurringRanRef = useRef(false);
  useEffect(() => {
    if (recurringRanRef.current) return;
    const rules = config.recurring || [];
    if (!rules.length) return;
    const { newTxs, updatedRecurring } = runRecurringRules(rules);
    if (newTxs.length > 0) {
      recurringRanRef.current = true;
      saveTxs([...newTxs, ...txs]);
      saveConfig({ ...config, recurring: updatedRecurring });
      showToast(`${newTxs.length} movimiento${newTxs.length === 1 ? "" : "s"} recurrente${newTxs.length === 1 ? "" : "s"} generado${newTxs.length === 1 ? "" : "s"}`);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [config.recurring]);

  const saveRecurring = (rules) => {
    // al guardar reglas, corre inmediatamente las que ya tengan fechas vencidas
    const { newTxs, updatedRecurring } = runRecurringRules(rules);
    if (newTxs.length > 0) {
      saveTxs([...newTxs, ...txs]);
      showToast(`${newTxs.length} movimiento${newTxs.length === 1 ? "" : "s"} recurrente${newTxs.length === 1 ? "" : "s"} generado${newTxs.length === 1 ? "" : "s"}`);
    }
    saveConfig({ ...config, recurring: updatedRecurring });
  };

  const upsertTx = (tx, learnedCats, linkInfo) => {
    const exists = txs.some((t) => t.id === tx.id);
    let nextTxs = exists ? txs.map((t) => (t.id === tx.id ? tx : t)) : [tx, ...txs];

    // Si nos pasaron información de vinculación, también marcamos a los demás
    // miembros del grupo como passThrough con el mismo groupId
    if (linkInfo && linkInfo.linkIds && linkInfo.linkIds.length && linkInfo.groupIdToUse) {
      const ids = new Set(linkInfo.linkIds);
      nextTxs = nextTxs.map((t) =>
        ids.has(t.id)
          ? { ...t, passThrough: true, groupId: linkInfo.groupIdToUse, categoryId: null }
          : t
      );
    }

    // Si dejamos de ser passThrough y antes estábamos en un grupo:
    // quitamos también del grupo a los huérfanos del grupo si quedan menos de 2
    if (exists && tx.passThrough === false) {
      const old = txs.find((t) => t.id === tx.id);
      if (old?.groupId) {
        const others = nextTxs.filter((t) => t.id !== tx.id && t.groupId === old.groupId);
        if (others.length <= 1) {
          // grupo de 1 ya no tiene sentido: quitar passThrough y groupId al solitario
          nextTxs = nextTxs.map((t) =>
            t.groupId === old.groupId && t.id !== tx.id
              ? { ...t, passThrough: false, groupId: null }
              : t
          );
        }
      }
    }

    saveTxs(nextTxs);
    if (learnedCats) saveConfig({ ...config, categories: learnedCats });
    setAdding(false);
    setEditingTx(null);
    showToast(exists ? "Movimiento actualizado" : `${tx.type === "income" ? "Ingreso" : "Gasto"} registrado`);
  };

  const saveManyTxs = (newTxs, learnedCats) => {
    saveTxs([...newTxs, ...txs]);
    if (learnedCats) saveConfig({ ...config, categories: learnedCats });
    setImportOpen(false);
    showToast(`${newTxs.length} movimiento${newTxs.length === 1 ? "" : "s"} importado${newTxs.length === 1 ? "" : "s"}`);
  };

  const overlayOpen = chatOpen || adding || !!editingTx || accountsOpen || importOpen || excelOpen || addMenuOpen || recurringOpen || settingsOpen;

  return (
    <div>
      <StickyHeader config={config} saveConfig={saveConfig} balance={balance} dateRange={dateRange} onOpenRange={() => setRangeOpen(true)} onOpenSettings={() => setSettingsOpen(true)} />

      <div className="cc-wrap">
        {tab === "inicio" && <Dashboard config={config} txs={txs} balance={balance} dateRange={dateRange} onEdit={setEditingTx} onAddAccount={() => setAccountsOpen(true)} saveConfig={saveConfig} />}
        {tab === "movs" && <Movimientos config={config} txs={txs} dateRange={dateRange} saveTxs={saveTxs} showToast={showToast} onEdit={setEditingTx} />}
        {tab === "cats" && <Categorias config={config} txs={txs} dateRange={dateRange} saveConfig={saveConfig} showToast={showToast} saveRecurring={saveRecurring} />}
        {tab === "stats" && <Estadisticas config={config} txs={txs} dateRange={dateRange} onEdit={setEditingTx} saveConfig={saveConfig} />}
      </div>

      <BottomNav
        tab={tab}
        setTab={setTab}
        onOpenAssistant={(opts) => { setChatVoice(!!(opts && opts.voice)); setChatOpen(true); }}
        hidden={overlayOpen}
      />

      <TopFab
        open={addMenuOpen}
        onToggle={() => setAddMenuOpen((v) => !v)}
        onPickExcel={() => { setAddMenuOpen(false); setExcelOpen(true); }}
        onPickScreenshot={() => { setAddMenuOpen(false); setImportOpen(true); }}
        onPickManual={() => { setAddMenuOpen(false); setAdding(true); }}
        onPickRecurring={() => { setAddMenuOpen(false); setRecurringOpen(true); }}
        hidden={chatOpen || adding || !!editingTx || accountsOpen || importOpen || excelOpen || recurringOpen}
      />

      {(adding || editingTx) && (
        <AddModal
          config={config}
          tx={editingTx}
          txs={txs}
          onClose={() => { setAdding(false); setEditingTx(null); }}
          onSave={upsertTx}
        />
      )}

      {importOpen && (
        <ImportModal
          config={config}
          txs={txs}
          onClose={() => setImportOpen(false)}
          onSave={saveManyTxs}
        />
      )}

      {chatOpen && (
        <Assistant
          config={config}
          txs={txs}
          saveConfig={saveConfig}
          saveTxs={saveTxs}
          autoVoice={chatVoice}
          onClose={() => { setChatOpen(false); setChatVoice(false); }}
          onOpenImport={() => setImportOpen(true)}
        />
      )}

      {accountsOpen && (
        <AccountsModal
          config={config}
          txs={txs}
          saveConfig={saveConfig}
          showToast={showToast}
          resetAll={resetAll}
          onClose={() => setAccountsOpen(false)}
        />
      )}

      {excelOpen && (
        <ExcelImportModal
          config={config}
          txs={txs}
          onClose={() => setExcelOpen(false)}
          onSave={saveManyTxs}
        />
      )}

      {recurringOpen && (
        <RecurringModal
          config={config}
          onClose={() => setRecurringOpen(false)}
          onSave={saveRecurring}
        />
      )}

      {settingsOpen && (
        <SettingsModal
          config={config}
          saveConfig={saveConfig}
          onClose={() => setSettingsOpen(false)}
          showToast={showToast}
          resetAll={resetAll}
        />
      )}

      {rangeOpen && (
        <DateRangeModal
          dateRange={dateRange}
          onClose={() => setRangeOpen(false)}
          onSave={(r) => { setDateRange(r); setRangeOpen(false); }}
        />
      )}
    </div>
  );
}

/* Iconos lineales del bottom nav (outline, sin emojis) */
const NavIconHome = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 11l9-8 9 8" />
    <path d="M5 10v10h14V10" />
  </svg>
);
const NavIconHistory = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 12a9 9 0 1 0 3-6.7" />
    <path d="M3 4v4h4" />
    <path d="M12 8v4l3 2" />
  </svg>
);
const NavIconCategories = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="3" width="7" height="7" rx="1.5" />
    <rect x="14" y="3" width="7" height="7" rx="1.5" />
    <rect x="3" y="14" width="7" height="7" rx="1.5" />
    <rect x="14" y="14" width="7" height="7" rx="1.5" />
  </svg>
);
const NavIconStats = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M4 20V10" />
    <path d="M12 20V4" />
    <path d="M20 20v-6" />
  </svg>
);

/* Iconos lineales para chips de cabecera (sin emojis) */
const IconGear = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <circle cx="12" cy="12" r="3" />
    <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.6 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.6a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
  </svg>
);
const IconCalendar = () => (
  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="4" width="18" height="18" rx="2" />
    <path d="M16 2v4M8 2v4M3 10h18" />
  </svg>
);

/* StickyHeader: header pegajoso con logo, balance, chip de rango y tabs */
/* Botón flotante del asistente — se monta en document.body para evitar
   problemas de stacking context (ancestros con transform/backdrop-filter
   rompen position:fixed). DOM nativo, no usa react-dom/createPortal. */
/* BottomNav: barra inferior con Home / History / Orb IA / Categories / Statistics */
function BottomNav({ tab, setTab, onOpenAssistant, hidden }) {
  const orbRef = useRef(null);
  const holdTimer = useRef(null);
  const heldRef = useRef(false);

  const orbBurst = () => {
    const orb = orbRef.current;
    if (!orb) return;
    const rect = orb.getBoundingClientRect();
    for (let i = 0; i < 6; i++) {
      const p = document.createElement("span");
      p.style.cssText = `position:fixed;width:4px;height:4px;border-radius:50%;background:var(--orb-mint);pointer-events:none;left:${rect.left + rect.width / 2}px;top:${rect.top + rect.height / 2}px;z-index:9999;`;
      document.body.appendChild(p);
      const angle = (i / 6) * Math.PI * 2;
      const dx = Math.cos(angle) * 22;
      const dy = Math.sin(angle) * 22;
      const anim = p.animate(
        [{ transform: "translate(0,0) scale(1)", opacity: 1 },
         { transform: `translate(${dx}px,${dy}px) scale(0)`, opacity: 0 }],
        { duration: 500, easing: "cubic-bezier(.2,.7,.2,1)" }
      );
      anim.onfinish = () => p.remove();
    }
    orb.animate(
      [{ transform: "translate(-50%,-50%) scale(1)" },
       { transform: "translate(-50%,-50%) scale(.8)", offset: .3 },
       { transform: "translate(-50%,-50%) scale(1.12)", offset: .6 },
       { transform: "translate(-50%,-50%) scale(1)" }],
      { duration: 550, easing: "cubic-bezier(.34,1.56,.64,1)" }
    );
  };

  // toque corto: abre asistente · mantener presionado: abre y dicta por voz
  const onOrbDown = () => {
    heldRef.current = false;
    if (!voiceSupported) return;
    holdTimer.current = setTimeout(() => {
      heldRef.current = true;
      if (navigator.vibrate) navigator.vibrate(12);
      orbBurst();
      onOpenAssistant({ voice: true });
    }, 450);
  };
  const onOrbUp = () => {
    if (holdTimer.current) { clearTimeout(holdTimer.current); holdTimer.current = null; }
  };
  const handleOrbClick = () => {
    if (heldRef.current) { heldRef.current = false; return; } // ya se abrió por hold
    orbBurst();
    setTimeout(() => onOpenAssistant({ voice: false }), 120);
  };

  const NAV_ITEMS = [
    ["inicio", "home",    NavIconHome],
    ["movs",   "history", NavIconHistory],
  ];
  const NAV_ITEMS_2 = [
    ["cats",  "categories", NavIconCategories],
    ["stats", "statistics", NavIconStats],
  ];

  return (
    <div className="cc-bottomnav" style={{
      opacity: hidden ? 0 : 1,
      transform: hidden ? "translateY(16px)" : "translateY(0)",
      pointerEvents: hidden ? "none" : "auto",
      transition: "opacity .2s, transform .25s cubic-bezier(.2,.7,.2,1)",
    }}>
      <div className="cc-bottomnav-inner">
        {NAV_ITEMS.map(([k, label, Icon]) => (
          <button key={k} className={`cc-nav-item ${tab === k ? "on" : ""}`} onClick={() => setTab(k)}>
            <span className="cc-nav-icon"><Icon /></span>
            <span className="cc-nav-label">{t(label)}</span>
          </button>
        ))}

        <div className="cc-orb-slot">
          <button ref={orbRef} className="cc-orb-btn"
            onClick={handleOrbClick}
            onMouseDown={onOrbDown} onMouseUp={onOrbUp} onMouseLeave={onOrbUp}
            onTouchStart={onOrbDown} onTouchEnd={onOrbUp}
            aria-label="Asistente Zafi (mantén presionado para hablar)">
            <span className="cc-orb" />
          </button>
        </div>

        {NAV_ITEMS_2.map(([k, label, Icon]) => (
          <button key={k} className={`cc-nav-item ${tab === k ? "on" : ""}`} onClick={() => setTab(k)}>
            <span className="cc-nav-icon"><Icon /></span>
            <span className="cc-nav-label">{t(label)}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

/* TopFab: botón circular (+) arriba a la derecha; abre hoja de captura */
function TopFab({ open, onToggle, onPickExcel, onPickScreenshot, onPickManual, onPickRecurring, hidden }) {
  const items = [
    { icon: "✏️", label: t("manualCapture"), desc: t("manualCaptureDesc"), onClick: onPickManual },
    { icon: "🔁", label: t("recurringMovement"), desc: t("recurringDesc"), onClick: onPickRecurring },
    { icon: "📸", label: t("fromScreenshot"), desc: t("fromScreenshotDesc"), onClick: onPickScreenshot },
    { icon: "📊", label: t("fromExcel"), desc: t("fromExcelDesc"), onClick: onPickExcel },
  ];
  return (
    <>
      <button className={`cc-fab-top ${open ? "open" : ""}`}
        onClick={onToggle}
        style={{
          opacity: hidden ? 0 : 1,
          transform: hidden ? `${open ? "rotate(45deg) " : ""}translateY(-16px)` : (open ? "rotate(45deg)" : "none"),
          pointerEvents: hidden ? "none" : "auto",
          transition: "opacity .2s, transform .25s cubic-bezier(.2,.7,.2,1)",
        }}
        aria-label="Nueva transacción">＋</button>

      {open && !hidden && (
        <div className="cc-overlay" onClick={onToggle}>
          <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
            <div className="cc-grip" />
            <div className="cc-sheet-top">
              <h2>{t("addMovement")}</h2>
              <button className="cc-sheet-close" onClick={onToggle}>×</button>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 9, marginTop: 4 }}>
              {items.map((it) => (
                <button key={it.label} className="cc-card" onClick={it.onClick}
                  style={{ display: "flex", alignItems: "center", gap: 14, padding: "15px 16px",
                    cursor: "pointer", textAlign: "left", width: "100%", fontFamily: "inherit" }}>
                  <div style={{ width: 44, height: 44, borderRadius: 12, background: "var(--surface)",
                    display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, flexShrink: 0 }}>
                    {it.icon}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: 15, color: "var(--ink)", letterSpacing: "-.01em" }}>{it.label}</div>
                    <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 2 }}>{it.desc}</div>
                  </div>
                  <span style={{ fontSize: 18, color: "var(--ink-faint)", flexShrink: 0 }}>›</span>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function StickyHeader({ config, saveConfig, balance, dateRange, onOpenRange, onOpenSettings }) {
  const [scrolled, setScrolled] = useState(false);
  useEffect(() => {
    const onScroll = () => setScrolled(window.scrollY > 4);
    window.addEventListener("scroll", onScroll, { passive: true });
    return () => window.removeEventListener("scroll", onScroll);
  }, []);
  // nombre del usuario — config.userName si existe, si no, derivado del correo
  const email = auth.currentUser?.email || "";
  const rawName = email.split("@")[0] || "Usuario";
  const emailName = rawName
    .replace(/[._\-]+/g, " ")
    .split(" ")
    .filter(Boolean)
    .map((w) => w[0].toUpperCase() + w.slice(1))
    .join(" ") || "Usuario";
  const displayName = config?.userName || emailName;
  const initial = displayName[0]?.toUpperCase() || "U";
  const avatarSrc = getAvatarSrc(config);

  return (
    <div className={`cc-top ${scrolled ? "scrolled" : ""}`}>
      <div className="cc-top-inner">
        {/* Logo centrado */}
        <div className="cc-zafi-wordmark">zafi</div>

        {/* Perfil: avatar + nombre + plan */}
        <div className="cc-profile-row">
          <button onClick={onOpenSettings}
            style={{ display: "flex", alignItems: "center", gap: 12, border: "none", background: "transparent",
              cursor: "pointer", padding: 0, textAlign: "left", minWidth: 0 }}>
            <div className="cc-avatar" style={{ overflow: "hidden" }}>
              {avatarSrc ? <img src={avatarSrc} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : initial}
            </div>
            <div style={{ minWidth: 0 }}>
              <div className="cc-profile-name">{displayName}</div>
              <div className="cc-profile-plan">Plan semanal</div>
            </div>
          </button>
          <div style={{ flex: 1 }} />
          <button className="cc-range-chip" onClick={onOpenRange}>
            <IconCalendar />
            <span>{rangeLabel(dateRange)}</span>
            <span className="cc-range-arrow">▼</span>
          </button>
        </div>
      </div>
    </div>
  );
}

/* ===================== AVATARES ========================================= */
const AVATAR_STYLES = [
  { id: "male-default",   url: "data:image/jpeg;base64,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" },
  { id: "female-default", url: "data:image/jpeg;base64,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" },
  { id: "option-beanie",  url: "data:image/jpeg;base64,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" },
  { id: "option-hoodie",  url: "data:image/jpeg;base64,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" },
];

// Asigna avatar por defecto según el género del registro
function defaultAvatarForGender(gender) {
  if (gender === "male") return "male-default";
  if (gender === "female") return "female-default";
  return null; // "other" → solo inicial
}

// Resizes an image file to a small square base64 for storage
function resizeImageToBase64(file, size = 200) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = size; canvas.height = size;
        const ctx = canvas.getContext("2d");
        const s = Math.min(img.width, img.height);
        const sx = (img.width - s) / 2, sy = (img.height - s) / 2;
        ctx.drawImage(img, sx, sy, s, s, 0, 0, size, size);
        resolve(canvas.toDataURL("image/jpeg", 0.8));
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

function getAvatarSrc(config) {
  if (config.avatarData) return config.avatarData; // custom photo
  if (config.avatarId) {
    const av = AVATAR_STYLES.find(a => a.id === config.avatarId);
    return av ? av.url : null;
  }
  return null;
}

function AvatarPickerModal({ config, saveConfig, onClose, showToast }) {
  const userName = config.userName || t("user");
  const currentSrc = getAvatarSrc(config);
  const initial = userName.charAt(0).toUpperCase();
  const fileRef = useRef(null);

  const pickPreset = (id) => {
    const { avatarData, ...rest } = config;
    saveConfig({ ...rest, avatarId: id });
    showToast(t("avatarUpdated"));
    onClose();
  };

  const handlePhoto = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const base64 = await resizeImageToBase64(file);
    const { avatarId, ...rest } = config;
    saveConfig({ ...rest, avatarData: base64 });
    showToast(t("avatarUpdated"));
    onClose();
  };

  const clear = () => {
    const { avatarId, avatarData, ...rest } = config;
    saveConfig(rest);
    showToast(t("avatarRemoved"));
    onClose();
  };

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{t("chooseAvatar")}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>

        {/* Current avatar */}
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8, marginBottom: 18 }}>
          <div style={{ position: "relative" }}>
            <div style={{ width: 80, height: 80, borderRadius: "50%", overflow: "hidden",
              background: "var(--ink)", display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 32, color: "#fff", fontFamily: "'Fraunces',serif", fontWeight: 600,
              boxShadow: "0 4px 20px rgba(30,40,60,.18)" }}>
              {currentSrc ? <img src={currentSrc} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : initial}
            </div>
            {(config.avatarId || config.avatarData) && (
              <button onClick={clear}
                style={{ position: "absolute", top: -4, right: -4, width: 26, height: 26,
                  borderRadius: "50%", border: "none", background: "var(--surface)",
                  boxShadow: "var(--shadow-sm)", cursor: "pointer", fontSize: 13,
                  display: "flex", alignItems: "center", justifyContent: "center", color: "var(--ink-soft)" }}>
                🗑
              </button>
            )}
          </div>
          <div style={{ fontWeight: 600, fontSize: 18, color: "var(--ink)" }}>{userName}</div>
        </div>

        {/* Upload photo buttons */}
        <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
          <button className="cc-btn" onClick={() => fileRef.current?.click()}
            style={{ flex: 1, padding: "12px 10px", fontSize: 13, display: "flex", flexDirection: "column", alignItems: "center", gap: 6 }}>
            <span style={{ fontSize: 22 }}>🖼</span>
            {t("choosePhoto")}
          </button>
          <input ref={fileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={handlePhoto} />
          <button className="cc-btn" onClick={() => {
            const inp = document.createElement("input");
            inp.type = "file"; inp.accept = "image/*"; inp.capture = "user";
            inp.onchange = handlePhoto; inp.click();
          }}
            style={{ flex: 1, padding: "12px 10px", fontSize: 13, display: "flex", flexDirection: "column", alignItems: "center", gap: 6 }}>
            <span style={{ fontSize: 22 }}>📷</span>
            {t("takePhoto")}
          </button>
        </div>

        {/* Preset avatars grid */}
        <div style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 11, fontWeight: 700,
          color: "var(--ink-faint)", letterSpacing: ".06em", textTransform: "uppercase",
          marginBottom: 10 }}>Avatares</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          {AVATAR_STYLES.map(a => (
            <button key={a.id} onClick={() => pickPreset(a.id)}
              style={{ aspectRatio: "1", borderRadius: 18, overflow: "hidden",
                border: config.avatarId === a.id ? "3px solid var(--green)" : "2px solid var(--glass-border)",
                cursor: "pointer", padding: 0,
                backgroundImage: `url(${a.url})`, backgroundSize: "cover", backgroundPosition: "center",
                backgroundColor: "var(--surface)",
                boxShadow: config.avatarId === a.id ? "0 0 0 2px var(--green)" : "var(--shadow-sm)",
                transition: "border .15s, box-shadow .15s" }}>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ===================== MODAL: AJUSTES GENERALES ========================= */
function SettingsModal({ config, saveConfig, onClose, showToast, resetAll }) {
  const user = auth.currentUser;
  const email = user?.email || "";
  const [userName, setUserName] = useState(config.userName || "");
  const [phone, setPhone] = useState(config.phone || "");
  const [lang, setLang] = useState(config.language || "es");
  const [currency, setCurrency] = useState(config.currency || "MXN");
  const [confirmReset, setConfirmReset] = useState(false);
  const [busy, setBusy] = useState(false);
  const [avatarOpen, setAvatarOpen] = useState(false);

  const initial = userName ? userName.charAt(0).toUpperCase() : email.charAt(0).toUpperCase();
  const avatarSrc = getAvatarSrc(config);

  const savePersonal = () => {
    saveConfig({ ...config, userName: userName.trim(), phone: phone.trim() });
    // update Firestore profile too
    if (user) {
      setDoc(doc(db, "users", user.uid, "data", "profile"), {
        name: userName.trim(), phone: phone.trim(), email,
      }, { merge: true }).catch(() => {});
    }
    showToast(t("infoUpdated"));
  };

  const saveLang = (l) => {
    setLang(l);
    setAppLang(l);
    saveConfig({ ...config, language: l });
    showToast(l === "es" ? "Idioma: Español" : "Language: English");
  };

  const saveCurrency = (c) => {
    setCurrency(c);
    saveConfig({ ...config, currency: c });
    showToast(`Moneda: ${c}`);
  };

  const doSignOut = async () => {
    setBusy(true);
    try { await signOut(auth); } catch (e) {}
    setBusy(false);
    onClose();
  };

  const doReset = () => {
    resetAll();
    setConfirmReset(false);
    onClose();
  };

  const ROW = { display: "flex", alignItems: "center", justifyContent: "space-between",
    padding: "14px 0", borderBottom: "1px solid var(--line-soft)" };
  const SECTION_TITLE = { fontFamily: "'Montserrat', sans-serif", fontSize: 11, fontWeight: 700,
    color: "var(--ink-faint)", letterSpacing: ".06em", textTransform: "uppercase",
    marginTop: 20, marginBottom: 6 };
  const ITEM_LABEL = { fontSize: 14, fontWeight: 500, color: "var(--ink)" };
  const ITEM_VALUE = { fontSize: 13, color: "var(--ink-soft)", textAlign: "right" };

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{t("settings")}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>

        {/* --- Perfil header --- */}
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 8, padding: "8px 0 18px",
          borderBottom: "1px solid var(--line-soft)" }}>
          <button onClick={() => setAvatarOpen(true)}
            style={{ position: "relative", border: "none", background: "transparent", cursor: "pointer", padding: 0 }}>
            <div className="cc-avatar" style={{ width: 72, height: 72, fontSize: 28, overflow: "hidden" }}>
              {avatarSrc ? <img src={avatarSrc} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} /> : initial}
            </div>
            <div style={{ position: "absolute", bottom: -2, right: -2, width: 24, height: 24, borderRadius: "50%",
              background: "var(--ink)", color: "#fff", fontSize: 12,
              display: "flex", alignItems: "center", justifyContent: "center",
              boxShadow: "var(--shadow-sm)" }}>✎</div>
          </button>
          <div style={{ fontWeight: 600, fontSize: 18, color: "var(--ink)" }}>{userName || "Usuario"}</div>
          <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>{email}</div>
        </div>

        {avatarOpen && (
          <AvatarPickerModal config={config} saveConfig={saveConfig} onClose={() => setAvatarOpen(false)} showToast={showToast} />
        )}

        {/* --- Información personal --- */}
        <div style={SECTION_TITLE}>{t("personalInfo")}</div>
        <div style={{ marginBottom: 6 }}>
          <label style={{ fontSize: 12, fontWeight: 600, color: "var(--ink-faint)", marginBottom: 4, display: "block" }}>Nombre</label>
          <input className="cc-input" value={userName} onChange={(e) => setUserName(e.target.value)}
            placeholder="Tu nombre" style={{ marginBottom: 10 }} />
          <label style={{ fontSize: 12, fontWeight: 600, color: "var(--ink-faint)", marginBottom: 4, display: "block" }}>Teléfono</label>
          <input className="cc-input" value={phone} onChange={(e) => setPhone(e.target.value)}
            placeholder="+52 664 123 4567" type="tel" style={{ marginBottom: 10 }} />
          <label style={{ fontSize: 12, fontWeight: 600, color: "var(--ink-faint)", marginBottom: 4, display: "block" }}>Correo electrónico</label>
          <input className="cc-input" value={email} disabled
            style={{ marginBottom: 10, opacity: 0.6 }} />
          <button className="cc-btn" style={{ fontSize: 13, padding: "8px 16px" }} onClick={savePersonal}>
            Guardar cambios
          </button>
        </div>

        {/* --- Idioma --- */}
        <div style={SECTION_TITLE}>{t("language")}</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 4 }}>
          {[["es", "🇲🇽 Español"], ["en", "🇺🇸 English"]].map(([k, l]) => (
            <button key={k} onClick={() => saveLang(k)}
              style={{ flex: 1, padding: "11px 8px", borderRadius: 12, cursor: "pointer",
                fontFamily: "'Montserrat', sans-serif", fontSize: 13.5,
                fontWeight: lang === k ? 600 : 400,
                background: lang === k ? "var(--ink)" : "var(--surface)",
                color: lang === k ? "#fff" : "var(--ink)",
                border: `1px solid ${lang === k ? "var(--ink)" : "var(--line)"}` }}>
              {l}
            </button>
          ))}
        </div>

        {/* --- Moneda --- */}
        <div style={SECTION_TITLE}>{t("currency")}</div>
        <div style={{ display: "flex", gap: 8, marginBottom: 4 }}>
          {[["MXN", "🇲🇽 MXN"], ["USD", "🇺🇸 USD"], ["EUR", "🇪🇺 EUR"]].map(([k, l]) => (
            <button key={k} onClick={() => saveCurrency(k)}
              style={{ flex: 1, padding: "11px 8px", borderRadius: 12, cursor: "pointer",
                fontFamily: "'Montserrat', sans-serif", fontSize: 13.5,
                fontWeight: currency === k ? 600 : 400,
                background: currency === k ? "var(--ink)" : "var(--surface)",
                color: currency === k ? "#fff" : "var(--ink)",
                border: `1px solid ${currency === k ? "var(--ink)" : "var(--line)"}` }}>
              {l}
            </button>
          ))}
        </div>

        {/* --- Notificaciones (placeholder) --- */}
        <div style={SECTION_TITLE}>{t("notifications")}</div>
        <div style={ROW}>
          <span style={ITEM_LABEL}>Recordatorios de gastos</span>
          <span style={ITEM_VALUE}>Próximamente</span>
        </div>

        {/* --- Datos --- */}
        <div style={SECTION_TITLE}>{t("dataPrivacy")}</div>
        <div style={ROW}>
          <span style={ITEM_LABEL}>{t("exportData")}</span>
          <span style={ITEM_VALUE}>Próximamente</span>
        </div>
        {!confirmReset ? (
          <button onClick={() => setConfirmReset(true)}
            style={{ ...ROW, width: "100%", border: "none", borderBottom: "1px solid var(--line-soft)",
              background: "transparent", cursor: "pointer", fontFamily: "inherit", textAlign: "left" }}>
            <span style={{ ...ITEM_LABEL, color: "var(--coral)" }}>Reiniciar app (borrar todo)</span>
          </button>
        ) : (
          <div style={{ padding: "12px 0", borderBottom: "1px solid var(--line-soft)" }}>
            <div style={{ fontSize: 13, color: "var(--coral)", fontWeight: 600, marginBottom: 8 }}>
              ¿Estás seguro? Esto borrará todas tus categorías, cuentas y movimientos.
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button className="cc-btn" style={{ flex: 1, padding: "8px 12px", fontSize: 13 }}
                onClick={() => setConfirmReset(false)}>Cancelar</button>
              <button className="cc-btn" style={{ flex: 1, padding: "8px 12px", fontSize: 13,
                background: "var(--coral)", color: "#fff", borderColor: "var(--coral)" }}
                onClick={doReset}>Sí, borrar todo</button>
            </div>
          </div>
        )}

        {/* --- Sesión --- */}
        <div style={{ marginTop: 20, marginBottom: 8 }}>
          <button className="cc-btn" onClick={doSignOut} disabled={busy}
            style={{ width: "100%", padding: 14, fontSize: 15, fontWeight: 600 }}>
            {busy ? t("signingOut") : "Cerrar sesión"}
          </button>
        </div>

        <div style={{ textAlign: "center", fontSize: 11, color: "var(--ink-faint)", padding: "6px 0 4px" }}>
          Zafi · Finanzas personales con IA · v1.0
        </div>
      </div>
    </div>
  );
}

/* ProfileNameModal: editar el nombre que se muestra en el header */
function ProfileNameModal({ current, onClose, onSave }) {
  const [name, setName] = useState(current || "");
  return (
    <div className="cc-overlay" onClick={onClose} style={{ alignItems: "center" }}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}
        style={{ borderRadius: 24, maxWidth: 360, width: "calc(100% - 40px)" }}>
        <div className="cc-grip" />
        <h2 className="cc-serif" style={{ fontSize: 20, fontWeight: 600, marginBottom: 14 }}>Tu nombre</h2>
        <input className="cc-input" placeholder="Ej. Luis Ángel" value={name}
          onChange={(e) => setName(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && name.trim() && onSave(name.trim())}
          style={{ marginBottom: 14 }} autoFocus />
        <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 13, fontSize: 14 }}
          disabled={!name.trim()} onClick={() => onSave(name.trim())}>
          Guardar
        </button>
      </div>
    </div>
  );
}

/* DateRangeModal: elegir el rango global de toda la app */
function DateRangeModal({ dateRange, onClose, onSave }) {
  const r = dateRange || DEFAULT_RANGE;
  const [preset, setPreset] = useState(r.preset);
  const resolved = resolveRange(r);
  const [from, setFrom] = useState(r.from || resolved.from);
  const [to, setTo] = useState(r.to || resolved.to);

  const PRESETS = [
    { id: "today",      label: "Hoy",              emoji: "⚡" },
    { id: "week",       label: "Esta semana",      emoji: "📅" },
    { id: "month",      label: "Este mes",         emoji: "📆" },
    { id: "last-month", label: "Mes pasado",       emoji: "📆" },
    { id: "3m",         label: "Últimos 3 meses",  emoji: "📊" },
    { id: "6m",         label: "Últimos 6 meses",  emoji: "📊" },
    { id: "year",       label: "Este año",         emoji: "🗓️" },
    { id: "last-year",  label: "Año pasado",       emoji: "🗓️" },
    { id: "all",        label: "Todo el historial", emoji: "⏳" },
  ];

  function apply() {
    if (preset === "custom") {
      if (!from || !to) return;
      const f = from > to ? to : from;
      const t = from > to ? from : to;
      onSave({ preset: "custom", from: f, to: t });
    } else {
      onSave({ preset });
    }
  }

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>Rango de tiempo</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 14 }}>
          Esta selección se aplica a todas las pestañas: Inicio, Movimientos, Categorías y Estadísticas.
        </p>

        <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 14 }}>
          {PRESETS.map((p) => (
            <button key={p.id} onClick={() => setPreset(p.id)}
              style={{ display: "flex", alignItems: "center", gap: 11, padding: "11px 13px",
                background: preset === p.id ? "var(--green-soft)" : "var(--surface)",
                border: `1px solid ${preset === p.id ? "var(--green)" : "var(--line)"}`,
                borderRadius: 11, cursor: "pointer", fontFamily: "inherit", fontSize: 14,
                fontWeight: preset === p.id ? 700 : 500, textAlign: "left", color: "var(--ink)" }}>
              <span style={{ fontSize: 17 }}>{p.emoji}</span>
              <span style={{ flex: 1 }}>{p.label}</span>
              {preset === p.id && <span style={{ color: "var(--green)", fontSize: 17, fontWeight: 700 }}>✓</span>}
            </button>
          ))}

          <button onClick={() => setPreset("custom")}
            style={{ display: "flex", alignItems: "center", gap: 11, padding: "11px 13px",
              background: preset === "custom" ? "var(--green-soft)" : "var(--surface)",
              border: `1px solid ${preset === "custom" ? "var(--green)" : "var(--line)"}`,
              borderRadius: 11, cursor: "pointer", fontFamily: "inherit", fontSize: 14,
              fontWeight: preset === "custom" ? 700 : 500, textAlign: "left", color: "var(--ink)",
              marginTop: 4 }}>
            <span style={{ fontSize: 17 }}>✏️</span>
            <span style={{ flex: 1 }}>Personalizado</span>
            {preset === "custom" && <span style={{ color: "var(--green)", fontSize: 17, fontWeight: 700 }}>✓</span>}
          </button>

          {preset === "custom" && (
            <div style={{ display: "flex", gap: 8, marginTop: 4, padding: "12px 13px",
              background: "var(--surface-2)", borderRadius: 11 }}>
              <div style={{ flex: 1 }}>
                <label className="cc-label" style={{ marginBottom: 4 }}>Desde</label>
                <input className="cc-input" type="date" value={from}
                  onChange={(e) => setFrom(e.target.value)} style={{ fontSize: 13 }} />
              </div>
              <div style={{ flex: 1 }}>
                <label className="cc-label" style={{ marginBottom: 4 }}>Hasta</label>
                <input className="cc-input" type="date" value={to}
                  onChange={(e) => setTo(e.target.value)} style={{ fontSize: 13 }} />
              </div>
            </div>
          )}
        </div>

        <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 13, fontSize: 14 }}
          onClick={apply}>
          Aplicar
        </button>
      </div>
    </div>
  );
}

/* secciones disponibles del inicio */
const DEFAULT_SECTIONS = [
  { id: "balance", label: "Saldo destacado", icon: "💰", on: false },
  { id: "kpis", label: "Ingresos y gastos del mes", icon: "📊", on: true },
  { id: "byCategory", label: "Gastos por categoría", icon: "🏷️", on: true },
  { id: "trend", label: "Mini gráfica de saldo (30d)", icon: "📈", on: true },
  { id: "topExpenses", label: "Gastos más grandes del mes", icon: "💸", on: false },
  { id: "recent", label: "Movimientos recientes", icon: "🕐", on: true },
];

function loadSections(config) {
  const saved = config.homeSections;
  if (!Array.isArray(saved) || !saved.length) return DEFAULT_SECTIONS;
  // sincronizar con DEFAULT por si agregamos nuevas
  const known = new Set(DEFAULT_SECTIONS.map((s) => s.id));
  const cleaned = saved.filter((s) => known.has(s.id));
  DEFAULT_SECTIONS.forEach((d) => {
    if (!cleaned.find((s) => s.id === d.id)) cleaned.push({ ...d });
  });
  return cleaned;
}

/* ============================= DASHBOARD ================================= */
function Dashboard({ config, txs, balance, dateRange, onEdit, onAddAccount, saveConfig }) {
  const [view, setView] = useState("all"); // "all" o accountId
  const [configuring, setConfiguring] = useState(false);
  const sections = loadSections(config);

  const scopedTxs = view === "all" ? txs : txs.filter((t) => t.accountId === view);
  // movimientos del rango global (en lugar de "mes actual")
  const rangeTxs = txsInRange(scopedTxs, dateRange);
  const monthStat = statTxs(rangeTxs).all;
  const inc = monthStat.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const exp = monthStat.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);

  const headerBalance = view === "all" ? balance : accountBalance(config, txs, view);
  const accName = view === "all" ? "General" : (config.accounts.find((a) => a.id === view) || {}).name || "";
  const headerLabel = view === "all" ? "Balance total" : `Saldo · ${accName}`;

  const byCat = {};
  monthStat.filter((t) => t.type === "expense" && t.categoryId).forEach((t) => {
    byCat[t.categoryId] = (byCat[t.categoryId] || 0) + t.amount;
  });
  const rows = Object.entries(byCat).sort((a, b) => b[1] - a[1]).slice(0, 6);
  const maxCat = rows.length ? rows[0][1] : 1;
  const catOf = (id) => config.categories.find((c) => c.id === id);

  // gastos más grandes del rango (excluyendo pass-through)
  const topExpenses = rangeTxs.filter((t) => t.type === "expense" && !t.passThrough)
    .sort((a, b) => b.amount - a.amount).slice(0, 5);

  // mini-gráfica de saldo (30 días)
  const trendPoints = (() => {
    const initial = view === "all"
      ? config.accounts.reduce((s, a) => s + (a.initialBalance || 0), 0)
      : (config.accounts.find((a) => a.id === view)?.initialBalance || 0);
    const start = new Date(); start.setDate(start.getDate() - 29);
    const startK = start.toISOString().slice(0, 10);
    const sorted = [...scopedTxs].sort((a, b) => a.date.localeCompare(b.date));
    let running = initial + sorted.filter((t) => t.date < startK).reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
    const map = new Map(); map.set(startK, running);
    sorted.forEach((t) => {
      if (t.date < startK) return;
      running += t.type === "income" ? t.amount : -t.amount;
      map.set(t.date, running);
    });
    const pts = [], todayD = new Date(today());
    let last = initial + sorted.filter((t) => t.date < startK).reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
    for (let d = new Date(start); d <= todayD; d.setDate(d.getDate() + 1)) {
      const k = d.toISOString().slice(0, 10);
      if (map.has(k)) last = map.get(k);
      pts.push({ date: k, val: last });
    }
    return pts;
  })();

  const isOn = (id) => sections.find((s) => s.id === id)?.on;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      {/* === selector de cuentas mejorado === */}
      {config.accounts.length > 0 && (
        <div className="cc-fade">
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
            <div className="cc-label">Tus cuentas</div>
            <button className="cc-gear" onClick={() => setConfiguring(true)}><IconGear /> Personalizar</button>
          </div>
          <div className="cc-scroll-x">
            {/* tarjeta General */}
            <button className={`cc-acc-card ${view === "all" ? "on" : ""}`} onClick={() => setView("all")}>
              <div className="cc-acc-icon">∑</div>
              <div className="cc-acc-label">Total</div>
              <div className="cc-acc-name">General</div>
              <div className="cc-acc-bal cc-num" style={{ color: balance < 0 ? "var(--coral)" : "var(--ink)" }}>
                {fmt(balance)}
              </div>
              <div className="cc-acc-sub">{config.accounts.length} cuenta{config.accounts.length === 1 ? "" : "s"}</div>
            </button>
            {/* tarjetas por cuenta */}
            {config.accounts.map((a) => {
              const b = accountBalance(config, txs, a.id);
              const accTxs = txs.filter((t) => t.accountId === a.id);
              const rangeAccStat = statTxs(txsInRange(accTxs, dateRange)).all;
              const rangeFlow = rangeAccStat.reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
              const active = view === a.id;
              return (
                <button key={a.id} className={`cc-acc-card ${active ? "on" : ""}`} onClick={() => setView(a.id)}>
                  <div className="cc-acc-icon">🏦</div>
                  <div className="cc-acc-label">Cuenta</div>
                  <div className="cc-acc-name">{a.name}</div>
                  <div className="cc-acc-bal cc-num" style={{ color: b < 0 ? "var(--coral)" : "var(--ink)" }}>
                    {fmt(b)}
                  </div>
                  <div className="cc-acc-sub" style={{ color: rangeFlow < 0 ? "var(--coral)" : undefined }}>
                    {rangeFlow >= 0 ? "▲" : "▼"} {fmt(Math.abs(rangeFlow))} en el periodo
                  </div>
                </button>
              );
            })}
            {/* tarjeta agregar */}
            <button className="cc-acc-card" onClick={onAddAccount}
              style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center",
                minWidth: 130, borderStyle: "dashed", color: "var(--ink-soft)" }}>
              <div style={{ fontSize: 30, lineHeight: 1, marginBottom: 6, color: "var(--gold)" }}>＋</div>
              <div style={{ fontWeight: 600, fontSize: 13 }}>Agregar cuenta</div>
            </button>
          </div>
        </div>
      )}

      {/* === secciones según orden y on/off === */}
      {sections.filter((s) => s.on).map((s, idx) => {
        const delay = `${idx * 60}ms`;

        if (s.id === "balance") return (
          <div key={s.id} className="cc-card" style={{ padding: "22px 22px" }}>
            <div className="cc-label">{headerLabel}</div>
            <div className="cc-serif cc-num" style={{ fontSize: 44, fontWeight: 600, letterSpacing: "-.02em", color: headerBalance < 0 ? "var(--coral)" : "var(--ink)" }}>
              {fmt(headerBalance)}
            </div>
          </div>
        );

        if (s.id === "kpis") return (
          <div key={s.id} style={{ display: "flex", gap: 10 }} className="cc-fade">
            <div className="cc-card" style={{ flex: 1, padding: 14 }}>
              <div className="cc-label">Ingresos · {rangeLabel(dateRange)}</div>
              <div className="cc-serif cc-num" style={{ fontSize: 21, fontWeight: 500, color: "var(--ink)" }}>{fmt(inc)}</div>
            </div>
            <div className="cc-card" style={{ flex: 1, padding: 14 }}>
              <div className="cc-label">Gastos · {rangeLabel(dateRange)}</div>
              <div className="cc-serif cc-num" style={{ fontSize: 21, fontWeight: 500, color: "var(--coral)" }}>{fmt(exp)}</div>
            </div>
          </div>
        );

        if (s.id === "byCategory") return (
          <div key={s.id} className="cc-card" style={{ padding: "6px 18px" }}>
            <div className="cc-label" style={{ marginTop: 12, marginBottom: 6 }}>Gastos por categoría · {rangeLabel(dateRange)}</div>
            {rows.length === 0 ? (
              <div style={{ color: "var(--ink-soft)", fontSize: 13, padding: "8px 0 14px" }}>
                No hay gastos en el periodo.
              </div>
            ) : (
              <div style={{ display: "flex", flexDirection: "column" }}>
                {rows.map(([id, amt]) => {
                  const c = catOf(id);
                  const pct = maxCat ? Math.round((amt / maxCat) * 100) : 0;
                  return (
                    <div key={id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "11px 0",
                      borderBottom: "1px solid var(--line-soft)" }}>
                      <div style={{ width: 34, height: 34, borderRadius: 10, background: "var(--surface)",
                        display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0 }}>
                        {c ? c.emoji : "❔"}
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 600, fontSize: 13, color: "var(--ink)", letterSpacing: "-.01em" }}>
                          {c ? c.name : "Sin categoría"}
                        </div>
                        <div style={{ height: 5, background: "var(--surface-2)", borderRadius: 99, overflow: "hidden", marginTop: 5 }}>
                          <div style={{ height: "100%", width: `${pct}%`, background: "var(--bar-fill)", borderRadius: 99 }} />
                        </div>
                      </div>
                      <div className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 14, color: "var(--coral)", whiteSpace: "nowrap" }}>
                        {fmtBare(amt)}<span style={{ fontSize: 10, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );

        if (s.id === "trend") return (
          <div key={s.id} className="cc-card" style={{ padding: 20 }}>
            <div className="cc-label" style={{ marginBottom: 8 }}>Saldo · últimos 30 días</div>
            {trendPoints.length < 2 ? (
              <div style={{ color: "var(--ink-soft)", fontSize: 14 }}>Datos insuficientes.</div>
            ) : (
              <LineChart points={trendPoints} />
            )}
          </div>
        );

        if (s.id === "topExpenses") return (
          <div key={s.id} className="cc-card" style={{ padding: 20 }}>
            <div className="cc-label" style={{ marginBottom: 10 }}>Gastos más grandes · {rangeLabel(dateRange)}</div>
            {topExpenses.length === 0 ? (
              <div style={{ color: "var(--ink-soft)", fontSize: 14 }}>Sin gastos en el periodo.</div>
            ) : (
              topExpenses.map((t) => <TxRow key={t.id} t={t} config={config} onEdit={onEdit} />)
            )}
          </div>
        );

        if (s.id === "recent") return (
          <div key={s.id} className="cc-card" style={{ padding: 20 }}>
            <div className="cc-label" style={{ marginBottom: 10 }}>
              Movimientos recientes{view !== "all" ? ` · ${accName}` : ""}
            </div>
            {scopedTxs.length === 0 ? (
              <div style={{ color: "var(--ink-soft)", fontSize: 14 }}>Sin movimientos todavía.</div>
            ) : (
              scopedTxs.slice(0, 5).map((t) => <TxRow key={t.id} t={t} config={config} onEdit={onEdit} />)
            )}
          </div>
        );

        return null;
      })}

      {configuring && (
        <HomeConfigModal
          sections={sections}
          onClose={() => setConfiguring(false)}
          onSave={(newSections) => {
            saveConfig({ ...config, homeSections: newSections });
            setConfiguring(false);
          }}
        />
      )}
    </div>
  );
}

/* ============= MODAL: PERSONALIZAR INICIO ============================== */
function HomeConfigModal({ sections, onClose, onSave }) {
  const [items, setItems] = useState(sections);
  const [dragIdx, setDragIdx] = useState(null);
  const [overIdx, setOverIdx] = useState(null);

  const toggle = (id) => {
    setItems((prev) => prev.map((s) => (s.id === id ? { ...s, on: !s.on } : s)));
  };

  // drag & drop con HTML5
  const onDragStart = (i) => (e) => {
    setDragIdx(i);
    e.dataTransfer.effectAllowed = "move";
    try { e.dataTransfer.setData("text/plain", String(i)); } catch (_) {}
  };
  const onDragOver = (i) => (e) => {
    e.preventDefault();
    if (i !== overIdx) setOverIdx(i);
  };
  const onDrop = (i) => (e) => {
    e.preventDefault();
    if (dragIdx == null || dragIdx === i) { setDragIdx(null); setOverIdx(null); return; }
    setItems((prev) => {
      const next = [...prev];
      const [moved] = next.splice(dragIdx, 1);
      next.splice(i, 0, moved);
      return next;
    });
    setDragIdx(null); setOverIdx(null);
  };
  const move = (i, dir) => {
    setItems((prev) => {
      const j = i + dir;
      if (j < 0 || j >= prev.length) return prev;
      const next = [...prev];
      [next[i], next[j]] = [next[j], next[i]];
      return next;
    });
  };
  const reset = () => setItems(DEFAULT_SECTIONS.map((s) => ({ ...s })));

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>Personalizar inicio</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 16 }}>
          Activa o desactiva secciones, y arrastra para reordenarlas.
        </p>

        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 18 }}>
          {items.map((s, i) => (
            <div key={s.id}
              draggable
              onDragStart={onDragStart(i)}
              onDragOver={onDragOver(i)}
              onDrop={onDrop(i)}
              onDragEnd={() => { setDragIdx(null); setOverIdx(null); }}
              className={`cc-sortable ${!s.on ? "disabled" : ""}`}
              style={{ borderColor: overIdx === i ? "var(--gold)" : "var(--line)",
                opacity: dragIdx === i ? 0.4 : (s.on ? 1 : 0.55) }}>
              <span className="cc-grip-h">⋮⋮</span>
              <span style={{ fontSize: 18 }}>{s.icon}</span>
              <span style={{ flex: 1, fontWeight: 600, fontSize: 14 }}>{s.label}</span>
              <button className="cc-btn" onClick={() => move(i, -1)} disabled={i === 0}
                style={{ padding: "4px 8px", fontSize: 12 }}>↑</button>
              <button className="cc-btn" onClick={() => move(i, 1)} disabled={i === items.length - 1}
                style={{ padding: "4px 8px", fontSize: 12 }}>↓</button>
              <label style={{ display: "inline-flex", alignItems: "center", cursor: "pointer", marginLeft: 4 }}>
                <input type="checkbox" checked={s.on} onChange={() => toggle(s.id)}
                  style={{ width: 18, height: 18, accentColor: "var(--green)" }} />
              </label>
            </div>
          ))}
        </div>

        <div style={{ display: "flex", gap: 8 }}>
          <button className="cc-btn" style={{ padding: 12 }} onClick={reset}>Restablecer</button>
          <button className="cc-btn cc-btn-green" style={{ flex: 1, padding: 13, fontSize: 14 }}
            onClick={() => onSave(items)}>
            Guardar cambios
          </button>
        </div>
      </div>
    </div>
  );
}

/* fila de transacción */
function TxRow({ t, config, onEdit, onDelete, selectable, selected, onToggle }) {
  const c = config.categories.find((x) => x.id === t.categoryId);
  const acc = config.accounts.find((a) => a.id === t.accountId);
  const multi = config.accounts.length > 1;

  // modo selección: toda la fila es checkbox
  if (selectable) {
    return (
      <div
        onClick={() => onToggle && onToggle(t.id)}
        style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 0",
          borderBottom: "1px solid var(--line)", cursor: "pointer",
          background: selected ? "var(--green-soft)" : "transparent",
          margin: "0 -10px", paddingLeft: 10, paddingRight: 10, borderRadius: selected ? 8 : 0 }}>
        <input type="checkbox" checked={!!selected} readOnly
          style={{ width: 19, height: 19, accentColor: "var(--green)" }} />
        <div className="cc-emoji" style={{ fontSize: 22, width: 28, textAlign: "center" }}>{c ? c.emoji : "❔"}</div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontWeight: 600, fontSize: 14.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
            {t.description || (c ? c.name : "Movimiento")}
          </div>
          <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>
            {c ? c.name : "Sin categoría"} · {t.date}{multi && acc ? ` · ${acc.name}` : ""}
          </div>
        </div>
        <div className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 15, color: t.type === "income" ? "var(--green)" : "var(--coral)", whiteSpace: "nowrap" }}>
          {t.type === "income" ? "+" : "−"}{fmtBare(t.amount).replace("-", "")}
          <span style={{ fontSize: 10.5, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
        </div>
      </div>
    );
  }

  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0",
      borderBottom: "1px solid var(--line-soft)" }}>
      {/* emoji pill */}
      <div
        onClick={() => onEdit && onEdit(t)}
        className="cc-emoji"
        style={{ width: 34, height: 34, borderRadius: 10, background: "var(--surface)",
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 17, flexShrink: 0, cursor: onEdit ? "pointer" : "default" }}>
        {c ? c.emoji : "❔"}
      </div>
      {/* text */}
      <div
        onClick={() => onEdit && onEdit(t)}
        style={{ flex: 1, minWidth: 0, cursor: onEdit ? "pointer" : "default" }}>
        <div style={{ fontWeight: 600, fontSize: 13, color: "var(--ink)",
          whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
          letterSpacing: "-.01em" }}>
          {t.description || (c ? c.name : "Movimiento")}
        </div>
        <div style={{ fontSize: 10.5, color: "var(--ink-soft)", marginTop: 1.5, fontWeight: 500 }}>
          {c ? c.name : "Sin categoría"}{multi && acc ? ` · ${acc.name}` : ""}
        </div>
      </div>
      {/* amount */}
      <div
        onClick={() => onEdit && onEdit(t)}
        className="cc-num"
        style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 15,
          color: t.type === "income" ? "var(--green)" : "var(--coral)",
          whiteSpace: "nowrap", letterSpacing: "-.01em",
          cursor: onEdit ? "pointer" : "default" }}>
        {t.type === "income" ? "+" : "−"}{fmtBare(t.amount).replace("-", "")}
        <span style={{ fontSize: 10.5, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
      </div>
      {/* delete × */}
      {onDelete && (
        <button className="cc-sheet-close"
          style={{ width: 28, height: 28, fontSize: 16, flexShrink: 0 }}
          onClick={(e) => { e.stopPropagation(); onDelete(t.id); }}>×</button>
      )}
    </div>
  );
}

/* agrupa la lista en [{type:'header', date, income, expense}, {type:'tx', t}, ...] */
function renderGroupedByDay(list) {
  const out = [];
  let currentDate = null, headerRef = null;
  for (const t of list) {
    if (t.date !== currentDate) {
      currentDate = t.date;
      headerRef = { type: "header", date: t.date, income: 0, expense: 0 };
      out.push(headerRef);
    }
    // los pass-through no inflan el total del día
    if (!t.passThrough) {
      if (t.type === "income") headerRef.income += t.amount;
      else headerRef.expense += t.amount;
    }
    out.push({ type: "tx", t });
  }
  return out;
}

/* ============================ MOVIMIENTOS ================================ */
function Movimientos({ config, txs, dateRange, saveTxs, showToast, onEdit }) {
  const [filter, setFilter] = useState("all");
  const [sortBy, setSortBy] = useState("date-desc"); // date-desc | date-asc | amount-desc | amount-asc | account
  const [selectMode, setSelectMode] = useState(false);
  const [selected, setSelected] = useState(new Set());
  const [confirmDelete, setConfirmDelete] = useState(null); // ids[] o null
  const [accView, setAccView] = useState("all"); // account filter
  const multi = config.accounts.length > 1;

  // filtrar por cuenta, luego por rango global, luego por tipo
  const accTxs = accView === "all" ? txs : txs.filter((t) => t.accountId === accView);
  const rangeTxs = txsInRange(accTxs, dateRange);
  const filtered = rangeTxs.filter((t) => filter === "all" || t.type === filter);
  const list = [...filtered].sort((a, b) => {
    if (sortBy === "date-desc") return b.date.localeCompare(a.date) || b.id.localeCompare(a.id);
    if (sortBy === "date-asc")  return a.date.localeCompare(b.date) || a.id.localeCompare(b.id);
    if (sortBy === "amount-desc") return b.amount - a.amount;
    if (sortBy === "amount-asc")  return a.amount - b.amount;
    if (sortBy === "account") {
      const an = (config.accounts.find((x) => x.id === a.accountId) || {}).name || "";
      const bn = (config.accounts.find((x) => x.id === b.accountId) || {}).name || "";
      return an.localeCompare(bn) || b.date.localeCompare(a.date);
    }
    return 0;
  });

  // resumen del periodo (excluyendo pass-through)
  const rangeStat = statTxs(rangeTxs).all;
  const totalIn = rangeStat.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const totalOut = rangeStat.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  // top 3 categorías del tipo filtrado (o de gastos para "Todos")
  const topCatType = filter === "income" ? "income" : "expense";
  const topByCat = {};
  rangeStat.filter((t) => t.type === topCatType && t.categoryId).forEach((t) => {
    topByCat[t.categoryId] = (topByCat[t.categoryId] || 0) + t.amount;
  });
  const topCatRows = Object.entries(topByCat).sort((a, b) => b[1] - a[1]).slice(0, 3);
  const topTotal = Object.values(topByCat).reduce((s, v) => s + v, 0);

  const toggleOne = (id) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };
  const selectAll = () => setSelected(new Set(list.map((t) => t.id)));
  const clearSel = () => setSelected(new Set());

  const askDeleteOne = (id) => setConfirmDelete([id]);
  const askDeleteMany = () => {
    if (selected.size === 0) return;
    setConfirmDelete(Array.from(selected));
  };
  const doDelete = () => {
    if (!confirmDelete) return;
    const set = new Set(confirmDelete);
    saveTxs(txs.filter((t) => !set.has(t.id)));
    showToast(`${confirmDelete.length} movimiento${confirmDelete.length === 1 ? "" : "s"} eliminado${confirmDelete.length === 1 ? "" : "s"}`);
    setConfirmDelete(null);
    setSelected(new Set());
    setSelectMode(false);
  };

  const exitSelect = () => { setSelectMode(false); setSelected(new Set()); };

  return (
    <div>
      {/* selector de cuenta */}
      {multi && (
        <div style={{ marginBottom: 12 }}>
          <div className="cc-scroll-x">
            <button className={`cc-acc-card ${accView === "all" ? "on" : ""}`} onClick={() => setAccView("all")}
              style={{ minWidth: 100 }}>
              <div className="cc-acc-label">Todas</div>
              <div className="cc-acc-name">General</div>
            </button>
            {config.accounts.map((a) => (
              <button key={a.id} className={`cc-acc-card ${accView === a.id ? "on" : ""}`} onClick={() => setAccView(a.id)}
                style={{ minWidth: 100 }}>
                <div className="cc-acc-label">🏦</div>
                <div className="cc-acc-name">{a.name}</div>
              </button>
            ))}
          </div>
        </div>
      )}
      {/* barra superior: filtro normal o info de selección */}
      {!selectMode ? (
        <>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
            <div className="cc-tabs" style={{ flex: 1 }}>
              {[["all", "Todos"], ["income", "Ingresos"], ["expense", "Gastos"]].map(([k, l]) => (
                <button key={k} className={`cc-tab ${filter === k ? "on" : ""}`} onClick={() => setFilter(k)}>{l}</button>
              ))}
            </div>
            {list.length > 0 && (
              <button className="cc-btn" style={{ padding: "8px 12px", fontSize: 12, whiteSpace: "nowrap" }}
                onClick={() => setSelectMode(true)}>
                Seleccionar
              </button>
            )}
          </div>
          {filtered.length > 0 && (
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
              <span style={{ fontSize: 12, fontWeight: 700, color: "var(--ink-soft)", textTransform: "uppercase", letterSpacing: ".06em" }}>
                Ordenar
              </span>
              <select className="cc-select" value={sortBy} onChange={(e) => setSortBy(e.target.value)}
                style={{ flex: 1, fontSize: 13, padding: "7px 10px" }}>
                <option value="date-desc">📅 Fecha — más reciente primero</option>
                <option value="date-asc">📅 Fecha — más antiguo primero</option>
                <option value="amount-desc">💰 Monto — mayor a menor</option>
                <option value="amount-asc">💰 Monto — menor a mayor</option>
                {config.accounts.length > 1 && (
                  <option value="account">🏦 Por cuenta</option>
                )}
              </select>
              <span style={{ fontSize: 11.5, color: "var(--ink-soft)", whiteSpace: "nowrap" }}>
                {filtered.length} mov.
              </span>
            </div>
          )}
        </>
      ) : (
        <div className="cc-card" style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, padding: "9px 12px",
          borderRadius: 14 }}>
          <button className="cc-sheet-close" onClick={exitSelect}
            style={{ width: 30, height: 30, fontSize: 17, flexShrink: 0 }}>×</button>
          <span style={{ flex: 1, fontWeight: 600, fontSize: 14, color: "var(--ink)" }}>
            {selected.size} seleccionado{selected.size === 1 ? "" : "s"}
          </span>
          <button className="cc-btn" style={{ padding: "6px 11px", fontSize: 12 }}
            onClick={selected.size === list.length ? clearSel : selectAll}>
            {selected.size === list.length ? "Ninguno" : "Todos"}
          </button>
          <button className="cc-btn" style={{ padding: "6px 11px", fontSize: 12,
            background: selected.size === 0 ? "var(--surface-2)" : "var(--coral)", color: selected.size === 0 ? "var(--ink-soft)" : "#fff",
            borderColor: selected.size === 0 ? "var(--line)" : "var(--coral)" }}
            disabled={selected.size === 0} onClick={askDeleteMany}>
            🗑 Eliminar
          </button>
        </div>
      )}

      {/* tarjeta resumen del periodo */}
      {!selectMode && rangeStat.length > 0 && (
        <SummaryCard filter={filter} totalIn={totalIn} totalOut={totalOut}
          topCatRows={topCatRows} topTotal={topTotal} config={config} />
      )}

      <div className="cc-card" style={{ padding: "6px 18px" }}>
        {list.length === 0 ? (
          <div style={{ color: "var(--ink-soft)", fontSize: 14, padding: "16px 0" }}>Nada por aquí todavía.</div>
        ) : (sortBy === "date-desc" || sortBy === "date-asc") ? (
          renderGroupedByDay(list).map((entry) =>
            entry.type === "header" ? (
              <div key={`h-${entry.date}`} className="cc-day-sep">
                {(() => { const p = dayParts(entry.date); return (
                  <>
                    <span className="cc-day-num">{p.num}</span>
                    <span className="cc-day-name">{p.desc}</span>
                  </>
                ); })()}
                <div className="cc-day-totals">
                  {entry.income > 0 && <span className="pos">+{fmtBare(entry.income)}</span>}
                  {entry.expense > 0 && <span className="neg">−{fmtBare(entry.expense)}</span>}
                </div>
              </div>
            ) : (
              <TxRow key={entry.t.id} t={entry.t} config={config}
                onEdit={selectMode ? null : onEdit}
                onDelete={selectMode ? null : askDeleteOne}
                selectable={selectMode}
                selected={selected.has(entry.t.id)}
                onToggle={toggleOne}
              />
            )
          )
        ) : (
          list.map((t) => (
            <TxRow key={t.id} t={t} config={config}
              onEdit={selectMode ? null : onEdit}
              onDelete={selectMode ? null : askDeleteOne}
              selectable={selectMode}
              selected={selected.has(t.id)}
              onToggle={toggleOne}
            />
          ))
        )}
      </div>

      {list.length > 0 && !selectMode && (
        <div style={{ fontSize: 12, color: "var(--ink-soft)", textAlign: "center", marginTop: 10 }}>
          Toca un movimiento para editarlo · ✕ para eliminarlo
        </div>
      )}

      {confirmDelete && (
        <ConfirmDialog
          title={confirmDelete.length === 1 ? "¿Eliminar este movimiento?" : `¿Eliminar ${confirmDelete.length} movimientos?`}
          message="Esta acción no se puede deshacer."
          confirmLabel="Eliminar"
          danger
          onCancel={() => setConfirmDelete(null)}
          onConfirm={doDelete}
        />
      )}
    </div>
  );
}

/* diálogo de confirmación reutilizable */
function ConfirmDialog({ title, message, confirmLabel = "Confirmar", danger, onCancel, onConfirm }) {
  return (
    <div className="cc-overlay" onClick={onCancel} style={{ alignItems: "center" }}>
      <div onClick={(e) => e.stopPropagation()}
        style={{ background: "var(--bg)", borderRadius: 18, maxWidth: 400, width: "90%", padding: 22, animation: "ccUp .25s" }}>
        <h3 className="cc-serif" style={{ fontSize: 19, fontWeight: 600, marginBottom: 8 }}>{title}</h3>
        {message && <p style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 18 }}>{message}</p>}
        <div style={{ display: "flex", gap: 8 }}>
          <button className="cc-btn" style={{ flex: 1, padding: 12 }} onClick={onCancel}>Cancelar</button>
          <button
            className={danger ? "cc-btn" : "cc-btn cc-btn-green"}
            style={{ flex: 1, padding: 12,
              background: danger ? "var(--coral)" : undefined,
              color: danger ? "#fff" : undefined,
              borderColor: danger ? "var(--coral)" : undefined }}
            onClick={onConfirm}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================ CATEGORÍAS ================================= */
function Categorias({ config, txs, dateRange, saveConfig, showToast, saveRecurring }) {
  const [editing, setEditing] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null); // categoría a eliminar
  const [recurringOpen, setRecurringOpen] = useState(false);
  const [accView, setAccView] = useState("all");
  const multi = config.accounts.length > 1;
  const visibleAccounts = accView === "all" ? config.accounts : config.accounts.filter((a) => a.id === accView);

  const save = (cat) => {
    let cats;
    if (cat.id) cats = config.categories.map((c) => (c.id === cat.id ? cat : c));
    else cats = [...config.categories, { ...cat, id: uid(), keywords: [] }];
    saveConfig({ ...config, categories: cats });
    setEditing(null);
    showToast("Categoría guardada");
  };
  const askDel = (cat) => setConfirmDel(cat);
  const doDel = () => {
    if (!confirmDel) return;
    saveConfig({ ...config, categories: config.categories.filter((c) => c.id !== confirmDel.id) });
    showToast("Categoría eliminada");
    setConfirmDel(null);
  };

  // calcular totales por categoría en el rango (excluyendo pass-through)
  const rangeStat = statTxs(txsInRange(txs || [], dateRange)).all;
  const totalsByCat = {};
  rangeStat.forEach((t) => {
    if (t.categoryId) totalsByCat[t.categoryId] = (totalsByCat[t.categoryId] || 0) + t.amount;
  });

  const recurring = config.recurring || [];
  const activeRec = recurring.filter((r) => isRecActive(r));
  const accName = (id) => config.accounts.find((a) => a.id === id)?.name || "—";
  const catFor = (id) => config.categories.find((c) => c.id === id);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ fontSize: 13, color: "var(--ink-soft)", padding: "0 6px" }}>
        Cada cuenta tiene sus propias categorías. Los totales son de <b>{rangeLabel(dateRange)}</b>.
      </div>

      {/* selector de cuenta */}
      {multi && (
        <div style={{ marginBottom: 4 }}>
          <div className="cc-scroll-x">
            <button className={`cc-acc-card ${accView === "all" ? "on" : ""}`} onClick={() => setAccView("all")}
              style={{ minWidth: 100 }}>
              <div className="cc-acc-label">Todas</div>
              <div className="cc-acc-name">General</div>
            </button>
            {config.accounts.map((a) => (
              <button key={a.id} className={`cc-acc-card ${accView === a.id ? "on" : ""}`} onClick={() => setAccView(a.id)}
                style={{ minWidth: 100 }}>
                <div className="cc-acc-label">🏦</div>
                <div className="cc-acc-name">{a.name}</div>
              </button>
            ))}
          </div>
        </div>
      )}

      {/* ===== Movimientos recurrentes ===== */}
      <div className="cc-card" style={{ padding: 18 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 4 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 18 }}>🔁</span>
            <span className="cc-serif" style={{ fontSize: 17, fontWeight: 600 }}>Movimientos recurrentes</span>
          </div>
          <button className="cc-btn" style={{ padding: "5px 11px", fontSize: 12 }} onClick={() => setRecurringOpen(true)}>
            {recurring.length ? "Gestionar" : "＋ Nuevo"}
          </button>
        </div>
        {recurring.length === 0 ? (
          <div style={{ fontSize: 13, color: "var(--ink-soft)", padding: "6px 0 2px" }}>
            No tienes movimientos recurrentes. Crea uno para que se registre automáticamente (renta, sueldo, suscripciones…).
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", marginTop: 6 }}>
            {recurring.map((r) => {
              const c = catFor(r.categoryId);
              return (
                <button key={r.id} onClick={() => setRecurringOpen(true)}
                  style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 0",
                    background: "transparent", border: "none", borderBottom: "1px solid var(--line-soft)",
                    cursor: "pointer", fontFamily: "inherit", textAlign: "left", width: "100%",
                    opacity: isRecActive(r) ? 1 : 0.5 }}>
                  <div className="cc-emoji" style={{ width: 34, height: 34, borderRadius: 10, background: "var(--surface)",
                    display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0 }}>
                    {c ? c.emoji : "🔁"}
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: 13.5, color: "var(--ink)", letterSpacing: "-.01em" }}>
                      {r.description}{!isRecActive(r) && <span style={{ fontWeight: 500, color: "var(--ink-faint)" }}> · pausado</span>}
                    </div>
                    <div style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 2 }}>
                      {FREQ_LABELS_FN()[r.freq]} · {accName(r.accountId)}{c ? ` · ${c.name}` : ""}
                    </div>
                  </div>
                  <div className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 14,
                    color: r.type === "income" ? "var(--green)" : "var(--coral)", whiteSpace: "nowrap" }}>
                    {r.type === "income" ? "+" : "−"}{fmtBare(r.amount)}<span style={{ fontSize: 10, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
                  </div>
                </button>
              );
            })}
            {activeRec.length > 0 && (
              <div style={{ fontSize: 11, color: "var(--ink-faint)", marginTop: 10 }}>
                {activeRec.length} activo{activeRec.length === 1 ? "" : "s"} · se generan solos en su fecha
              </div>
            )}
          </div>
        )}
      </div>

      {visibleAccounts.map((acc) => {
        const accCats = config.categories.filter((c) => c.accountId === acc.id);
        return (
          <div key={acc.id} className="cc-card" style={{ padding: 18 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
              <span style={{ fontSize: 18 }}>🏦</span>
              <span className="cc-serif" style={{ fontSize: 17, fontWeight: 600 }}>{acc.name}</span>
            </div>
            {[["expense", "Gastos"], ["income", "Ingresos"]].map(([type, label]) => {
              const list = accCats.filter((c) => c.type === type);
              return (
                <div key={type}>
                  <div className="cc-label" style={{ marginTop: 12, marginBottom: 4 }}>{label}</div>
                  {list.length === 0 && (
                    <div style={{ fontSize: 13, color: "var(--ink-soft)", padding: "3px 0" }}>Sin categorías.</div>
                  )}
                  {list.map((c) => {
                    const total = totalsByCat[c.id] || 0;
                    return (
                      <div key={c.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 0", borderBottom: "1px solid var(--line)" }}>
                        <span className="cc-emoji" style={{ fontSize: 19 }}>{c.emoji}</span>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 600, fontSize: 14 }}>{c.name}</div>
                          {total > 0 && (
                            <div className="cc-num" style={{ fontSize: 11.5, color: "var(--ink-soft)" }}>
                              {fmt(total)} en el periodo
                            </div>
                          )}
                        </div>
                        <button className="cc-btn" style={{ padding: "4px 9px", fontSize: 12 }} onClick={() => setEditing(c)}>Editar</button>
                        <button className="cc-btn" style={{ padding: "4px 8px", fontSize: 12 }} onClick={() => askDel(c)}>✕</button>
                      </div>
                    );
                  })}
                </div>
              );
            })}
            <button className="cc-btn" style={{ marginTop: 12, fontSize: 13 }}
              onClick={() => setEditing({ accountId: acc.id, type: "expense", emoji: "📦", name: "" })}>
              ＋ Agregar categoría a {acc.name}
            </button>
          </div>
        );
      })}
      {editing && <CatModal cat={editing} accounts={config.accounts} onClose={() => setEditing(null)} onSave={save} />}
      {recurringOpen && (
        <RecurringModal config={config} onClose={() => setRecurringOpen(false)} onSave={saveRecurring} />
      )}
      {confirmDel && (
        <ConfirmDialog
          title="¿Eliminar esta categoría?"
          message={<>La categoría <b>{confirmDel.emoji} {confirmDel.name}</b> se eliminará. Los movimientos asociados quedarán sin categoría (no se borran).</>}
          confirmLabel="Eliminar"
          danger
          onCancel={() => setConfirmDel(null)}
          onConfirm={doDel}
        />
      )}
    </div>
  );
}

function CatModal({ cat, accounts, onClose, onSave }) {
  const [name, setName] = useState(cat.name || "");
  const [emoji, setEmoji] = useState(cat.emoji || "📦");
  const [type, setType] = useState(cat.type || "expense");
  const [accountId, setAccountId] = useState(cat.accountId || accounts[0].id);

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{cat.id ? "Editar categoría" : "Nueva categoría"}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <div style={{ width: 76 }}>
            <label className="cc-label">Emoji</label>
            <input className="cc-input" style={{ textAlign: "center", fontSize: 22 }} value={emoji}
              onChange={(e) => setEmoji(e.target.value.slice(0, 2) || "📦")} />
          </div>
          <div style={{ flex: 1 }}>
            <label className="cc-label">Nombre</label>
            <input className="cc-input" value={name} placeholder="Ej. Mascotas" onChange={(e) => setName(e.target.value)} />
          </div>
        </div>
        {accounts.length > 1 && (
          <div style={{ marginBottom: 14 }}>
            <label className="cc-label">Cuenta</label>
            <select className="cc-select" value={accountId} onChange={(e) => setAccountId(e.target.value)}>
              {accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </div>
        )}
        <label className="cc-label">Tipo</label>
        <div className="cc-tabs" style={{ marginBottom: 22 }}>
          {[["expense", "Gasto"], ["income", "Ingreso"]].map(([k, l]) => (
            <button key={k} className={`cc-tab ${type === k ? "on" : ""}`} onClick={() => setType(k)}>{l}</button>
          ))}
        </div>
        <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 13 }}
          disabled={!name.trim()}
          onClick={() => onSave({ ...cat, name: name.trim(), emoji, type, accountId })}>
          Guardar
        </button>
      </div>
    </div>
  );
}

/* ================== MODAL: GESTIONAR CUENTAS ============================= */
function AccountsModal({ config, txs, saveConfig, showToast, resetAll, onClose }) {
  const [editing, setEditing] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [confirmReset, setConfirmReset] = useState(false);

  const save = (acc) => {
    let accounts, categories = config.categories;
    if (acc.id) {
      accounts = config.accounts.map((a) => (a.id === acc.id ? acc : a));
    } else {
      const nid = uid();
      accounts = [...config.accounts, { ...acc, id: nid }];
      categories = [...config.categories, ...defaultCatsForAccount(nid)];
    }
    saveConfig({ ...config, accounts, categories, accountMode: accounts.length > 1 ? "multiple" : "single" });
    setEditing(null);
    showToast("Cuenta guardada");
  };
  const askDel = (acc) => {
    if (config.accounts.length <= 1) {
      showToast("Esta es tu última cuenta. Si quieres reiniciar, usa 'Empezar desde cero' abajo");
      return;
    }
    setConfirmDel(acc);
  };
  const doDel = () => {
    if (!confirmDel) return;
    const id = confirmDel.id;
    const accounts = config.accounts.filter((a) => a.id !== id);
    const categories = config.categories.filter((c) => c.accountId !== id);
    saveConfig({ ...config, accounts, categories, accountMode: accounts.length > 1 ? "multiple" : "single" });
    showToast("Cuenta eliminada");
    setConfirmDel(null);
  };
  const doReset = async () => {
    setConfirmReset(false);
    onClose();
    if (resetAll) await resetAll();
  };

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>Tus cuentas</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <div className="cc-card" style={{ padding: 14, marginBottom: 12 }}>
          {config.accounts.map((a) => {
            const b = accountBalance(config, txs, a.id);
            return (
              <div key={a.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "11px 0", borderBottom: "1px solid var(--line)" }}>
                <span style={{ fontSize: 20 }}>🏦</span>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 600, fontSize: 15 }}>{a.name}</div>
                  <div style={{ fontSize: 11.5, color: "var(--ink-soft)" }}>Saldo inicial: {fmt(a.initialBalance || 0)}</div>
                </div>
                <span className="cc-num" style={{ fontWeight: 700, color: b < 0 ? "var(--coral)" : "var(--green)" }}>{fmt(b)}</span>
                <button className="cc-btn" style={{ padding: "5px 10px", fontSize: 12 }} onClick={() => setEditing(a)}>Editar</button>
                <button className="cc-btn" style={{ padding: "5px 9px", fontSize: 12 }} onClick={() => askDel(a)}>✕</button>
              </div>
            );
          })}
          <button className="cc-btn cc-btn-primary" style={{ marginTop: 14, fontSize: 13, width: "100%" }}
            onClick={() => setEditing({ name: "", initialBalance: 0 })}>
            ＋ Agregar cuenta
          </button>
        </div>
        <div style={{ fontSize: 13, color: "var(--ink-soft)", padding: "0 6px" }}>
          El <b>saldo inicial</b> es el dinero que ya tienes en cada cuenta. El saldo de la derecha es ese monto ajustado con tus movimientos.
        </div>

        {/* Cerrar sesión */}
        <div style={{ marginTop:22, paddingTop:18, borderTop:"1px solid var(--line)" }}>
          <button onClick={() => signOut(auth)}
            style={{ width:"100%", padding:"11px 14px", fontSize:14, fontFamily:"inherit",
              background:"var(--surface-2)", color:"var(--ink-soft)",
              border:"1px solid var(--line)", borderRadius:12, cursor:"pointer",
              display:"flex", alignItems:"center", justifyContent:"center", gap:8 }}>
            🚪 Cerrar sesión
          </button>
        </div>

        {/* Zona peligrosa: empezar desde cero */}
        {resetAll && (
          <div style={{ marginTop: 22, paddingTop: 18, borderTop: "1px dashed var(--line)" }}>
            <div className="cc-label" style={{ color: "var(--coral)", marginBottom: 8 }}>Zona peligrosa</div>
            <button
              onClick={() => setConfirmReset(true)}
              style={{ width: "100%", padding: "11px 14px", fontFamily: "inherit", fontSize: 13.5, fontWeight: 600,
                background: "var(--coral-soft)", color: "var(--coral)",
                border: "1px solid var(--coral)", borderRadius: 12, cursor: "pointer",
                transition: ".2s", display: "flex", alignItems: "center", justifyContent: "center", gap: 7 }}>
              🔄 Empezar desde cero
            </button>
            <div style={{ fontSize: 11.5, color: "var(--ink-soft)", marginTop: 7, lineHeight: 1.4 }}>
              Borra todas tus cuentas, categorías y movimientos. La app vuelve al onboarding inicial como usuario nuevo. No se puede deshacer.
            </div>
          </div>
        )}

        {editing && <AccountModal acc={editing} onClose={() => setEditing(null)} onSave={save} />}
        {confirmDel && (
          <ConfirmDialog
            title="¿Eliminar esta cuenta?"
            message={<>La cuenta <b>{confirmDel.name}</b> y todas sus categorías se eliminarán. Los movimientos en esa cuenta quedarán huérfanos.</>}
            confirmLabel="Eliminar"
            danger
            onCancel={() => setConfirmDel(null)}
            onConfirm={doDel}
          />
        )}
        {confirmReset && (
          <ConfirmDialog
            title="¿Empezar desde cero?"
            message={<>Esto eliminará <b>TODAS</b> tus cuentas, categorías y movimientos para siempre. La app volverá al onboarding como si fueras un usuario nuevo. <b>Esta acción no se puede deshacer.</b></>}
            confirmLabel="Sí, borrar todo"
            danger
            onCancel={() => setConfirmReset(false)}
            onConfirm={doReset}
          />
        )}
      </div>
    </div>
  );
}

function AccountModal({ acc, onClose, onSave }) {
  const [name, setName] = useState(acc.name || "");
  const [bal, setBal] = useState(acc.initialBalance != null ? String(acc.initialBalance) : "");

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{acc.id ? "Editar cuenta" : "Nueva cuenta"}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <div style={{ marginBottom: 14 }}>
          <label className="cc-label">Nombre</label>
          <input className="cc-input" placeholder="Efectivo, BBVA, Santander, Tarjeta…" value={name}
            onChange={(e) => setName(e.target.value)} />
        </div>
        <div style={{ marginBottom: 22 }}>
          <label className="cc-label">Saldo inicial</label>
          <div className="cc-amount-display">
            <span className="cc-amount-currency">$</span>
            <input className="cc-num" type="text" inputMode="decimal" placeholder="0.00"
              value={bal}
              onChange={(e) => {
                const v = e.target.value.replace(/[^0-9.]/g, "");
                if ((v.match(/\./g) || []).length <= 1) setBal(v);
              }}
              style={{ width: `${Math.max((bal || "0.00").length, 4)}ch` }} />
            <span className="cc-amount-mxn">mxn</span>
          </div>
          <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 6, textAlign: "center" }}>
            Dinero que tienes ahorita en esta cuenta. Puede ser 0.
          </div>
        </div>
        <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 13 }}
          disabled={!name.trim()}
          onClick={() => onSave({ ...acc, name: name.trim(), initialBalance: parseFloat(bal) || 0 })}>
          Guardar
        </button>
      </div>
    </div>
  );
}

/* ===================== MODAL: NUEVO MOVIMIENTO =========================== */
function AddModal({ config, tx, txs, onClose, onSave }) {
  const editing = !!tx;
  const [type, setType] = useState(tx ? tx.type : "expense");
  const [amount, setAmount] = useState(tx ? String(tx.amount) : "");
  const [desc, setDesc] = useState(tx ? tx.description : "");
  const [accountId, setAccountId] = useState(
    tx ? tx.accountId : (config.accounts.length === 1 ? config.accounts[0].id : "")
  );
  const [date, setDate] = useState(tx ? tx.date : today());
  const [catId, setCatId] = useState(tx ? tx.categoryId : "auto");
  const [phase, setPhase] = useState("form");

  const cats = config.categories.filter((c) => c.type === type && c.accountId === accountId);

  /* categorizador: keywords locales -> Claude -> preguntar */
  async function categorize() {
    const nd = norm(desc);
    // primero, claves locales de las categorías de ESTA cuenta
    for (const c of cats) {
      for (const kw of c.keywords || []) {
        if (nd.includes(norm(kw))) return { catId: c.id, sure: true };
      }
    }
    // si no hay match, revisar claves aprendidas en otras cuentas con el mismo nombre/tipo
    for (const c of cats) {
      const twins = config.categories.filter(
        (x) => x.id !== c.id && x.type === c.type && norm(x.name).trim() === norm(c.name).trim()
      );
      for (const twin of twins) {
        for (const kw of twin.keywords || []) {
          if (nd.includes(norm(kw))) return { catId: c.id, sure: true };
        }
      }
    }
    try {
      const sys =
        'Eres un clasificador de transacciones personales. Responde SOLO con un objeto JSON, sin markdown ni texto extra: {"categoryId":"<id>" o null,"confident":true|false}. Devuelve null y confident:false si no estás razonablemente seguro.';
      const user = `Tipo: ${type === "income" ? "ingreso" : "gasto"}. Descripción: "${desc}". Categorías disponibles: ${JSON.stringify(
        cats.map((c) => ({ id: c.id, name: c.name }))
      )}`;
      const raw = await callClaude(sys, [{ role: "user", content: user }]);
      const p = parseJSON(raw);
      const valid = cats.some((c) => c.id === p.categoryId);
      if (valid && p.confident) return { catId: p.categoryId, sure: true };
      if (valid) return { catId: p.categoryId, sure: false };
    } catch (e) { /* sin API: preguntar */ }
    return { catId: null, sure: false };
  }

  function finalize(finalCatId, learn) {
    let learnedCats = null;
    if (learn && finalCatId) {
      const kws = extractKW(desc).slice(0, 3);
      if (kws.length) {
        const target = config.categories.find((c) => c.id === finalCatId);
        const sameName = target ? norm(target.name).trim() : null;
        learnedCats = config.categories.map((c) =>
          sameName && norm(c.name).trim() === sameName && c.type === target.type
            ? { ...c, keywords: [...new Set([...(c.keywords || []), ...kws])] }
            : c
        );
      }
    }
    onSave(
      { id: tx ? tx.id : uid(), type, amount: Math.abs(parseFloat(amount)),
        description: desc.trim(), categoryId: finalCatId, accountId, date },
      learnedCats,
      null
    );
  }

  async function handleSave() {
    if (!amount || parseFloat(amount) <= 0) return;
    if (catId !== "auto") { finalize(catId, true); return; }
    // Si solo hay una categoría disponible, usarla directo
    if (cats.length === 1) { finalize(cats[0].id, true); return; }
    // Intentar detectar con keywords locales primero (sin IA)
    const nd = norm(desc);
    for (const c of cats) {
      for (const kw of c.keywords || []) {
        if (nd.includes(norm(kw))) { finalize(c.id, false); return; }
      }
    }
    // Si hay desc, intentar con IA — pero si falla, ir directo a preguntar
    if (desc.trim()) {
      setPhase("detecting");
      const r = await categorize();
      if (r.catId && r.sure) { finalize(r.catId, false); return; }
      if (r.catId) setCatId(r.catId);
    }
    // Sin match seguro → preguntar al usuario
    setPhase("ask");
  }

  /* pantalla: preguntar categoría */
  if (phase === "ask") {
    return (
      <div className="cc-overlay" onClick={onClose}>
        <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
          <div className="cc-grip" />
          <h2 className="cc-serif" style={{ fontSize: 21, fontWeight: 600, marginBottom: 6 }}>¿Qué categoría es?</h2>
          <p style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 16 }}>
            No estoy seguro de la categoría para «{desc || "este movimiento"}». Elígela y la app aprenderá para la próxima.
          </p>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 9 }}>
            {cats.map((c) => (
              <button key={c.id} type="button"
                onClick={() => { finalize(c.id, true); }}
                style={{ padding: "13px 12px", textAlign: "left", cursor: "pointer", fontSize: 14, fontWeight: 600,
                  display: "flex", alignItems: "center", gap: 8, fontFamily: "inherit",
                  background: "var(--paper)", border: "1px solid var(--line)", borderRadius: 14,
                  boxShadow: "var(--shadow-xs)", WebkitTapHighlightColor: "transparent",
                  position: "relative", zIndex: 1 }}>
                <span className="cc-emoji" style={{ fontSize: 19 }}>{c.emoji}</span>{c.name}
              </button>
            ))}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{editing ? "Editar transacción" : "Nueva transacción"}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>

        <div className="cc-tabs" style={{ marginBottom: 4 }}>
          {[["expense", "Gasto"], ["income", "Ingreso"]].map(([k, l]) => (
            <button key={k} className={`cc-tab ${type === k ? "on" : ""}`}
              onClick={() => { setType(k); setCatId("auto"); }}>{l}</button>
          ))}
        </div>

        <div className="cc-amount-display">
          <span className="cc-amount-currency">$</span>
          <input className="cc-num" type="text" inputMode="decimal" placeholder="0.00"
            value={amount}
            onChange={(e) => {
              const v = e.target.value.replace(/[^0-9.]/g, "");
              if ((v.match(/\./g) || []).length <= 1) setAmount(v);
            }}
            style={{ width: `${Math.max((amount || "0.00").length, 4)}ch` }} />
          <span className="cc-amount-mxn">mxn</span>
        </div>

        <div style={{ marginBottom: 14 }}>
          <label className="cc-label">Concepto</label>
          <input className="cc-input" placeholder="Ej. tacos con la familia, gasolina, pago de luz…"
            value={desc} onChange={(e) => setDesc(e.target.value)} />
        </div>

        <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <div style={{ flex: 1 }}>
            <label className="cc-label">Fecha</label>
            <input className="cc-input" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
          </div>
          <div style={{ flex: 1 }}>
            <label className="cc-label">Categoría</label>
            <select className="cc-select" value={catId} onChange={(e) => setCatId(e.target.value)}
              disabled={!accountId}>
              <option value="auto">✨ Detectar automáticamente</option>
              {cats.map((c) => <option key={c.id} value={c.id}>{c.emoji} {c.name}</option>)}
            </select>
          </div>
        </div>

        {config.accounts.length > 1 && (
          <div style={{ marginBottom: 14 }}>
            <label className="cc-label">Cuenta</label>
            <select className="cc-select" value={accountId}
              onChange={(e) => { setAccountId(e.target.value); setCatId("auto"); }}
              style={{ borderColor: !accountId ? "var(--gold)" : "var(--line)" }}>
              <option value="">Elegir cuenta…</option>
              {config.accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </div>
        )}

        <div style={{ marginBottom: 20 }}>
          {!accountId ? (
            <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>
              Elige primero una cuenta para ver sus categorías.
            </div>
          ) : catId === "auto" && (
            <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>
              La app intentará detectar la categoría con el concepto. Si no está segura, te pregunta.
            </div>
          )}
        </div>

        <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 14, fontSize: 15 }}
          disabled={phase === "detecting" || !amount || parseFloat(amount) <= 0 || !accountId}
          onClick={handleSave}>
          {phase === "detecting"
            ? "Detectando categoría…"
            : editing ? "Guardar cambios" : `Guardar ${type === "income" ? "ingreso" : "gasto"}`}
        </button>
      </div>
    </div>
  );
}

/* ===================== ASISTENTE DE CHAT ================================= */
// El historial del chat persiste en memoria mientras la app esté abierta,
// para que no se pierda al cerrar y reabrir el asistente.
let CHAT_MSGS_STORE = null;
let CHAT_HISTORY_STORE = [];
function Assistant({ config, txs, saveConfig, saveTxs, onClose, onOpenImport, autoVoice }) {
  const GREET =
    "¡Hola! Soy tu asistente. Dime qué quieres y yo lo hago en la app — crear o quitar categorías, cuentas, registrar movimientos… o pregúntame sobre tus gastos.";
  const QUICK = [
    "Crea la categoría Mascotas 🐶",
    "Registra un gasto de 250 en gasolina",
    "¿Cuánto llevo gastado este mes?",
  ];
  const [msgs, setMsgs] = useState(() => CHAT_MSGS_STORE || [{ role: "bot", text: GREET }]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const history = useRef(CHAT_HISTORY_STORE);
  const scroller = useRef(null);

  // mantener el store sincronizado para que persista al cerrar/reabrir
  useEffect(() => { CHAT_MSGS_STORE = msgs; }, [msgs]);

  // ===== Dictado por voz =====
  const [listening, setListening] = useState(false);
  const listeningRef = useRef(false);
  const recRef = useRef(null);
  const finalRef = useRef("");
  const lastHeardRef = useRef("");
  const sendOnEndRef = useRef(false);

  const setListen = (v) => { listeningRef.current = v; setListening(v); };

  const startVoice = () => {
    if (!voiceSupported || busy || listeningRef.current) return;
    try {
      const rec = new SpeechRec();
      rec.lang = "es-MX";
      rec.continuous = true;
      rec.interimResults = true;
      finalRef.current = "";
      lastHeardRef.current = "";
      rec.onresult = (e) => {
        let interim = "";
        let final = "";
        for (let i = e.resultIndex; i < e.results.length; i++) {
          const t = e.results[i][0].transcript;
          if (e.results[i].isFinal) final += t;
          else interim += t;
        }
        if (final) finalRef.current += final;
        const full = (finalRef.current + interim).trim();
        lastHeardRef.current = full; // captura también lo provisional
        setInput(full);
      };
      rec.onerror = () => { setListen(false); };
      rec.onend = () => {
        setListen(false);
        // al terminar (tras soltar) ya están todos los resultados: enviar
        if (sendOnEndRef.current) {
          sendOnEndRef.current = false;
          const text = (lastHeardRef.current || finalRef.current).trim();
          if (text) send(text);
        }
      };
      recRef.current = rec;
      rec.start();
      setListen(true);
      if (navigator.vibrate) navigator.vibrate(8);
    } catch (err) { setListen(false); }
  };

  const stopVoice = (autoSend = true) => {
    if (!listeningRef.current) return;
    sendOnEndRef.current = autoSend; // el envío ocurre en onend (cuando ya hay texto final)
    try { recRef.current && recRef.current.stop(); } catch (e) {}
    if (navigator.vibrate) navigator.vibrate(8);
  };

  // Si se abrió manteniendo presionado el orbe, empieza a escuchar y
  // termina cuando el usuario suelte (en cualquier parte de la pantalla).
  useEffect(() => {
    if (!autoVoice || !voiceSupported) return;
    const t = setTimeout(() => startVoice(), 250);
    const release = () => stopVoice(true);
    window.addEventListener("mouseup", release);
    window.addEventListener("touchend", release);
    return () => {
      clearTimeout(t);
      window.removeEventListener("mouseup", release);
      window.removeEventListener("touchend", release);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (scroller.current) scroller.current.scrollTop = scroller.current.scrollHeight;
  }, [msgs, busy]);

  async function send(text) {
    const userText = (text ?? input).trim();
    if (!userText || busy) return;
    setMsgs((m) => [...m, { role: "me", text: userText }]);
    setInput("");
    setBusy(true);
    history.current.push({ role: "user", content: userText });
    try {
      const raw = await callClaude(assistantSystem(config, txs), history.current);
      const parsed = parseJSON(raw);
      history.current.push({ role: "assistant", content: parsed.message || raw });
      let log = [], charts = [];
      if (Array.isArray(parsed.actions) && parsed.actions.length) {
        // separar show_chart de las demás acciones
        const chartSpecs = parsed.actions.filter((a) => a && a.type === "show_chart");
        const otherActions = parsed.actions.filter((a) => a && a.type !== "show_chart");

        if (otherActions.length) {
          const res = applyActions(config, txs, otherActions);
          saveConfig(res.config);
          saveTxs(res.txs);
          log = res.log;
        }
        for (const cs of chartSpecs) {
          const built = buildChart(cs.spec || cs, config, txs);
          if (built) charts.push(built);
        }
      }
      setMsgs((m) => [...m, { role: "bot", text: parsed.message || "Listo.", log, charts }]);
    } catch (e) {
      setMsgs((m) => [
        ...m,
        { role: "bot", text: "Tuve un problema para procesar eso. ¿Me lo dices de otra forma?" },
      ]);
    }
    setBusy(false);
  }

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
          <h2 className="cc-serif" style={{ fontSize: 21, fontWeight: 600, display: "flex", alignItems: "center", gap: 9 }}>
            <span style={{ width: 10, height: 10, borderRadius: "50%", background: "var(--green)", boxShadow: "0 0 0 4px var(--green-soft)" }} />
            Asistente
          </h2>
          <div style={{ display: "flex", gap: 7 }}>
            {msgs.length > 1 && (
              <button className="cc-btn" style={{ padding: "6px 10px", fontSize: 13 }}
                onClick={() => {
                  setMsgs([{ role: "bot", text: GREET }]);
                  history.current.length = 0;
                  CHAT_MSGS_STORE = null;
                }}
                title="Limpiar conversación">🗑</button>
            )}
            <button className="cc-btn" style={{ padding: "6px 12px", fontSize: 13 }} onClick={onClose}>{t("close")}</button>
          </div>
        </div>

        <div ref={scroller} style={{ display: "flex", flexDirection: "column", gap: 10, overflowY: "auto", maxHeight: "52vh", padding: "2px 2px 6px" }}>
          {msgs.map((m, i) => (
            <div key={i} style={{ display: "flex", flexDirection: "column", gap: 8, alignItems: m.role === "me" ? "flex-end" : "stretch" }}>
              <div className={`cc-bubble ${m.role}`} style={{ whiteSpace: "pre-wrap" }}>
                {m.text}
                {m.log && m.log.length > 0 && (
                  <div style={{ marginTop: 9, display: "flex", flexDirection: "column", gap: 5 }}>
                    {m.log.map((l, j) => (
                      <div key={j} style={{
                        fontSize: 12.5, fontWeight: 600, padding: "5px 9px", borderRadius: 8,
                        background: l.ok ? "rgba(26,122,110,.14)" : "rgba(181,69,58,.12)",
                        color: l.ok ? "var(--green)" : "var(--coral)",
                      }}>
                        {l.ok ? "✓" : "!"} {l.text}
                      </div>
                    ))}
                  </div>
                )}
              </div>
              {/* gráficas fuera de la burbuja, ancho completo */}
              {m.charts && m.charts.length > 0 && m.charts.map((ch, j) => (
                <MiniChart key={j} data={ch} />
              ))}
            </div>
          ))}
          {busy && (
            <div className="cc-bubble bot"><span className="cc-dots"><span /><span /><span /></span></div>
          )}
        </div>

        {msgs.length === 1 && !busy && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 7, margin: "12px 0 2px" }}>
            {QUICK.map((q, i) => (
              <button key={i} className="cc-chip" onClick={() => send(q)}>{q}</button>
            ))}
          </div>
        )}

        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
          <button className="cc-btn" title="Importar desde screenshot"
            onClick={() => { onClose(); onOpenImport && onOpenImport(); }}
            style={{ padding: "10px 12px", fontSize: 18, lineHeight: 1 }}>📸</button>
          {voiceSupported && (
            <button
              className={`cc-btn cc-mic ${listening ? "rec" : ""}`}
              title="Mantén presionado para hablar"
              disabled={busy}
              onMouseDown={(e) => { e.preventDefault(); startVoice(); }}
              onMouseUp={() => stopVoice(true)}
              onMouseLeave={() => listening && stopVoice(true)}
              onTouchStart={(e) => { e.preventDefault(); startVoice(); }}
              onTouchEnd={(e) => { e.preventDefault(); stopVoice(true); }}
              style={{ padding: "10px 13px", fontSize: 18, lineHeight: 1, userSelect: "none", touchAction: "none" }}>
              {listening ? "●" : "🎙"}
            </button>
          )}
          <input
            className="cc-input"
            placeholder={listening ? "Escuchando…" : "Dile algo… ej. quita la categoría Ropa"}
            value={input}
            disabled={busy}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && send()}
          />
          <button className="cc-btn cc-btn-green" disabled={busy || !input.trim()} onClick={() => send()}>
            Enviar
          </button>
        </div>
        {listening && (
          <div style={{ fontSize: 11.5, color: "var(--ink-soft)", textAlign: "center", marginTop: 8 }}>
            Suelta para enviar
          </div>
        )}
      </div>
    </div>
  );
}

/* ===================== MODAL: MOVIMIENTOS RECURRENTES =================== */
function RecurringModal({ config, onClose, onSave }) {
  const rules = config.recurring || [];
  const [view, setView] = useState(rules.length ? "list" : "form"); // list | form
  const [editingId, setEditingId] = useState(null);

  // form state
  const [type, setType] = useState("expense");
  const [amount, setAmount] = useState("");
  const [desc, setDesc] = useState("");
  const [accountId, setAccountId] = useState(config.accounts.length === 1 ? config.accounts[0].id : "");
  const [catId, setCatId] = useState("auto");
  const [freq, setFreq] = useState("monthly");
  const [startDate, setStartDate] = useState(today());
  const [detecting, setDetecting] = useState(false);

  const cats = config.categories.filter((c) => c.type === type && c.accountId === accountId);

  const resetForm = () => {
    setType("expense"); setAmount(""); setDesc("");
    setAccountId(config.accounts.length === 1 ? config.accounts[0].id : "");
    setCatId("auto"); setFreq("monthly"); setStartDate(today()); setEditingId(null);
  };

  const startNew = () => { resetForm(); setView("form"); };

  const startEdit = (r) => {
    setType(r.type); setAmount(String(r.amount)); setDesc(r.description);
    setAccountId(r.accountId); setCatId(r.categoryId || "auto");
    setFreq(r.freq); setStartDate(r.startDate); setEditingId(r.id);
    setView("form");
  };

  const canSave = amount && parseFloat(amount) > 0 && desc.trim() && accountId;

  const save = async () => {
    if (!canSave) return;
    // detección automática de categoría si está en "auto"
    let finalCat = catId === "auto" ? null : (catId || null);
    if (catId === "auto") {
      setDetecting(true);
      try {
        const res = await autoCategorize(desc.trim(), type, accountId, config);
        finalCat = res.catId;
      } catch (e) { finalCat = null; }
      setDetecting(false);
    }
    const rule = {
      id: editingId || uid(),
      type, amount: Math.abs(parseFloat(amount)),
      description: desc.trim(), accountId,
      categoryId: finalCat,
      freq, startDate,
      lastRun: editingId ? (rules.find((r) => r.id === editingId)?.lastRun ?? null) : null,
      active: true,
      paused: false,
    };
    const next = editingId
      ? rules.map((r) => (r.id === editingId ? rule : r))
      : [...rules, rule];
    onSave(next);
    onClose();
  };

  const toggleActive = (id) => {
    onSave(rules.map((r) => {
      if (r.id !== id) return r;
      const nowActive = !isRecActive(r);
      return { ...r, active: nowActive, paused: !nowActive };
    }));
  };

  const remove = (id) => {
    onSave(rules.filter((r) => r.id !== id));
  };

  const accName = (id) => config.accounts.find((a) => a.id === id)?.name || "—";
  const catEmoji = (id) => config.categories.find((c) => c.id === id)?.emoji || "🔁";

  // ===== Vista LISTA =====
  if (view === "list") {
    return (
      <div className="cc-overlay" onClick={onClose}>
        <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
          <div className="cc-grip" />
          <div className="cc-sheet-top">
            <h2>{t("recurringMovements")}</h2>
            <button className="cc-sheet-close" onClick={onClose}>×</button>
          </div>
          <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 16 }}>
            Se generan automáticamente en su fecha. Puedes editarlos como cualquier movimiento.
          </p>

          {rules.length === 0 ? (
            <div style={{ color: "var(--ink-soft)", fontSize: 14, padding: "8px 0 18px" }}>
              Aún no tienes movimientos recurrentes.
            </div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 16 }}>
              {rules.map((r) => (
                <div key={r.id} className="cc-card" style={{ padding: "12px 14px", opacity: isRecActive(r) ? 1 : 0.5 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <div className="cc-emoji" style={{ width: 36, height: 36, borderRadius: 10, background: "var(--surface)",
                      display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, flexShrink: 0 }}>
                      {catEmoji(r.categoryId)}
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: 14, color: "var(--ink)", letterSpacing: "-.01em" }}>{r.description}</div>
                      <div style={{ fontSize: 11.5, color: "var(--ink-soft)", marginTop: 2 }}>
                        {FREQ_LABELS_FN()[r.freq]} · {accName(r.accountId)}
                      </div>
                    </div>
                    <div className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 14,
                      color: r.type === "income" ? "var(--green)" : "var(--coral)", whiteSpace: "nowrap" }}>
                      {r.type === "income" ? "+" : "−"}{fmtBare(r.amount)}
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 7, marginTop: 10 }}>
                    <button className="cc-btn" style={{ flex: 1, padding: "6px 10px", fontSize: 12 }}
                      onClick={() => startEdit(r)}>Editar</button>
                    <button className="cc-btn" style={{ flex: 1, padding: "6px 10px", fontSize: 12 }}
                      onClick={() => toggleActive(r.id)}>{isRecActive(r) ? "Pausar" : "Activar"}</button>
                    <button className="cc-btn" style={{ padding: "6px 10px", fontSize: 12, color: "var(--coral)" }}
                      onClick={() => remove(r.id)}>Eliminar</button>
                  </div>
                </div>
              ))}
            </div>
          )}

          <button className="cc-btn cc-btn-primary" style={{ width: "100%", padding: 13 }}
            onClick={startNew}>＋ Nuevo recurrente</button>
        </div>
      </div>
    );
  }

  // ===== Vista FORM =====
  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{editingId ? "Editar recurrente" : "Nuevo recurrente"}</h2>
          <button className="cc-sheet-close" onClick={() => (rules.length ? setView("list") : onClose())}>×</button>
        </div>

        <div className="cc-tabs" style={{ marginBottom: 4 }}>
          {[["expense", "Gasto"], ["income", "Ingreso"]].map(([k, l]) => (
            <button key={k} className={`cc-tab ${type === k ? "on" : ""}`}
              onClick={() => { setType(k); setCatId(""); }}>{l}</button>
          ))}
        </div>

        <div className="cc-amount-display">
          <span className="cc-amount-currency">$</span>
          <input className="cc-num" type="text" inputMode="decimal" placeholder="0.00"
            value={amount}
            onChange={(e) => {
              const v = e.target.value.replace(/[^0-9.]/g, "");
              if ((v.match(/\./g) || []).length <= 1) setAmount(v);
            }}
            style={{ width: `${Math.max((amount || "0.00").length, 4)}ch` }} />
          <span className="cc-amount-mxn">mxn</span>
        </div>

        <div style={{ marginBottom: 14 }}>
          <label className="cc-label">Concepto</label>
          <input className="cc-input" placeholder="Ej. renta, Netflix, sueldo…"
            value={desc} onChange={(e) => setDesc(e.target.value)} />
        </div>

        <div style={{ marginBottom: 14 }}>
          <label className="cc-label">¿Cada cuánto?</label>
          <div style={{ display: "flex", gap: 7, flexWrap: "wrap" }}>
            {Object.entries(FREQ_LABELS_FN()).map(([k, l]) => (
              <button key={k} onClick={() => setFreq(k)}
                style={{ padding: "9px 14px", borderRadius: 12, cursor: "pointer", fontFamily: "inherit",
                  fontSize: 13, fontWeight: freq === k ? 700 : 500,
                  background: freq === k ? "var(--ink)" : "var(--surface)",
                  color: freq === k ? "#fff" : "var(--ink)",
                  border: `1px solid ${freq === k ? "var(--ink)" : "var(--line)"}` }}>
                {l}
              </button>
            ))}
          </div>
        </div>

        <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <div style={{ flex: 1 }}>
            <label className="cc-label">Empieza el</label>
            <input className="cc-input" type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} />
          </div>
          {config.accounts.length > 1 && (
            <div style={{ flex: 1 }}>
              <label className="cc-label">Cuenta</label>
              <select className="cc-select" value={accountId} onChange={(e) => { setAccountId(e.target.value); setCatId(""); }}>
                <option value="">Elegir…</option>
                {config.accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
            </div>
          )}
        </div>

        <div style={{ marginBottom: 18 }}>
          <label className="cc-label">Categoría {accountId ? "" : "(elige cuenta primero)"}</label>
          <select className="cc-select" value={catId} onChange={(e) => setCatId(e.target.value)} disabled={!accountId}>
            <option value="auto">✨ Detectar automáticamente</option>
            <option value="">Sin categoría</option>
            {cats.map((c) => <option key={c.id} value={c.id}>{c.emoji} {c.name}</option>)}
          </select>
          {catId === "auto" && (
            <div style={{ fontSize: 11.5, color: "var(--ink-soft)", marginTop: 6 }}>
              Zafi elegirá la categoría según el concepto al guardar.
            </div>
          )}
        </div>

        <div style={{ fontSize: 12, color: "var(--ink-soft)", marginBottom: 14, lineHeight: 1.5 }}>
          Se generará un {type === "income" ? "ingreso" : "gasto"} de <b>{amount ? fmtBare(parseFloat(amount) || 0) : "$0"}</b> {FREQ_LABELS_FN()[freq].toLowerCase()} a partir del {startDate}. Si la fecha de inicio ya pasó, se crearán los movimientos faltantes al guardar.
        </div>

        <button className="cc-btn cc-btn-primary" style={{ width: "100%", padding: 14, opacity: (canSave && !detecting) ? 1 : 0.4 }}
          disabled={!canSave || detecting} onClick={save}>
          {detecting ? "Detectando categoría…" : (editingId ? "Guardar cambios" : "Crear recurrente")}
        </button>
      </div>
    </div>
  );
}

/* ===================== MODAL: IMPORTAR DESDE SCREENSHOTS ================= */
function ImportModal({ config, txs, onClose, onSave }) {
  const [files, setFiles] = useState([]); // {file, previewUrl}
  const [defaultAccountId, setDefaultAccountId] = useState(
    config.accounts.length === 1 ? config.accounts[0].id : ""
  );
  const [phase, setPhase] = useState("upload"); // upload | processing | review
  const [drafts, setDrafts] = useState([]); // movimientos extraídos editables
  const [error, setError] = useState(null);
  const inputRef = useRef(null);

  const addFiles = (list) => {
    const imgs = Array.from(list || []).filter((f) => f.type.startsWith("image/"));
    const mapped = imgs.map((f) => ({ file: f, previewUrl: URL.createObjectURL(f) }));
    setFiles((prev) => [...prev, ...mapped]);
  };

  const removeFile = (idx) => {
    setFiles((prev) => {
      URL.revokeObjectURL(prev[idx].previewUrl);
      return prev.filter((_, i) => i !== idx);
    });
  };

  async function process() {
    if (!files.length || !defaultAccountId) return;
    setError(null);
    setPhase("processing");
    try {
      const imagesB64 = await Promise.all(files.map((f) => fileToB64(f.file)));
      const acc = config.accounts.find((a) => a.id === defaultAccountId);
      const cats = config.categories
        .filter((c) => c.accountId === defaultAccountId)
        .map((c) => ({ name: c.name, type: c.type, emoji: c.emoji }));

      const sys = `Eres un extractor de movimientos bancarios mexicanos. Vas a recibir una o varias imágenes (screenshots de estados de cuenta o apps bancarias) y debes extraer TODOS los movimientos visibles.

Cuenta destino: ${acc ? acc.name : "Principal"}
Categorías disponibles para esa cuenta:
${JSON.stringify(cats)}

Responde SIEMPRE con UN SOLO objeto JSON sin markdown:
{"movements":[{"date":"YYYY-MM-DD","amount":<número positivo>,"type":"income"|"expense","description":"comercio o concepto","categoryName":"<nombre de categoría existente, o null si no estás seguro>"},...]}

REGLAS GENERALES:
- "date" en formato YYYY-MM-DD. Si la imagen solo muestra día/mes, usa el año actual (${today().slice(0,4)}). Si no hay fecha, usa la de hoy: ${today()}.
- "amount" SIEMPRE positivo, sin signo.
- "type": "expense" si es un cargo/gasto/débito; "income" si es un abono/depósito/crédito.
- "description": el comercio o concepto que aparece, breve (ej. "OXXO", "Uber", "Pago tarjeta", "Spotify").
- "categoryName": elige una de las categorías de arriba que coincida con el tipo (gasto/ingreso). Si no estás seguro, usa null.
- Ignora encabezados, saldos totales, comisiones que no sean movimientos reales.
- Si la imagen no muestra movimientos, devuelve {"movements":[]}.

🚨 CÓMO ASIGNAR LA FECHA — REGLA CRÍTICA, LEE CON CUIDADO:
Muchas apps bancarias mexicanas (Santander, BBVA, Banorte) muestran los separadores de fecha como un ENCABEZADO que aplica a los movimientos que aparecen DEBAJO de ese separador, NO a los de arriba.

Ejemplo del formato típico de Santander:
  [movimiento A] ←  ¿de qué día es?
  ─── sábado 02 de mayo 2026 ───
  [movimiento B]
  [movimiento C]
  ─── viernes 01 de mayo 2026 ───
  [movimiento D]

En este ejemplo:
- B y C son del SÁBADO 02 de mayo (debajo de su separador).
- D es del VIERNES 01 de mayo.
- A NO se le puede asignar fecha con certeza desde esta sola imagen porque su separador está cortado/fuera de pantalla (probablemente sea de un día más reciente, como lunes 04).

REGLAS DE FECHA:
1. La fecha de cada movimiento es la del separador que aparece JUSTO ARRIBA de él, o el más cercano hacia arriba.
2. Si un movimiento aparece en la parte SUPERIOR de la imagen sin un separador de fecha visible arriba de él (porque la captura empieza en medio del scroll), NO INVENTES su fecha. En ese caso, omite ese movimiento — aparecerá completo en otro screenshot del lote.
3. Las fechas dentro de la app van de más recientes (arriba) a más antiguas (abajo). Si ves un separador "sábado 02 mayo" y arriba hay movimientos sin separador visible, esos movimientos son de un día POSTERIOR al 02 (lunes 04, martes 05, etc.), no anterior.

🚨 EVITA DUPLICADOS ENTRE IMÁGENES:
Cuando subo varios screenshots de la misma app, los movimientos se traslapan entre fotos. Si detectas el mismo movimiento (mismo monto, mismo comercio) en dos imágenes distintas, inclúyelo UNA SOLA VEZ con la fecha correcta de su separador. Combina la información de ambas imágenes: si en una imagen el movimiento aparece cortado en la parte superior (sin fecha visible) y en otra imagen aparece bajo un separador claro, usa la fecha del separador.

- Incluye TODOS los movimientos que puedas datar con certeza, en el orden en que aparecen.`;

      const raw = await callClaudeVision(sys, "Extrae todos los movimientos visibles.", imagesB64);
      const parsed = parseJSON(raw);
      const list = Array.isArray(parsed.movements) ? parsed.movements : [];
      if (!list.length) {
        setError("No detecté movimientos en las imágenes. Verifica que se vean claros.");
        setPhase("upload");
        return;
      }
      // mapear a drafts con resolución de categoría y detección de duplicados
      const ds = list.map((m, i) => {
        const c = m.categoryName ? findCat({ ...config }, m.categoryName, defaultAccountId) : null;
        const candidate = {
          date: m.date || today(),
          amount: Math.abs(Number(m.amount) || 0),
          type: m.type === "income" ? "income" : "expense",
          description: m.description || "",
          accountId: defaultAccountId,
        };
        const dup = findDuplicate(candidate, txs || []);
        return {
          tempId: "draft" + i,
          ...candidate,
          categoryId: c ? c.id : "",
          selected: !dup, // desmarcar duplicados por defecto
          duplicate: dup ? { date: dup.date, amount: dup.amount, description: dup.description } : null,
        };
      });
      setDrafts(ds);
      setPhase("review");
    } catch (e) {
      console.error(e);
      setError("Hubo un problema procesando las imágenes. Inténtalo de nuevo.");
      setPhase("upload");
    }
  }

  function updateDraft(tempId, patch) {
    setDrafts((prev) => prev.map((d) => (d.tempId === tempId ? { ...d, ...patch } : d)));
  }

  function saveAll() {
    const valid = drafts.filter((d) => d.selected && d.amount > 0);
    if (!valid.length) return;
    // aprender de cada categoría asignada manualmente
    let cats = config.categories;
    valid.forEach((d) => {
      if (!d.categoryId || !d.description) return;
      const target = cats.find((c) => c.id === d.categoryId);
      if (!target) return;
      const kws = extractKW(d.description).slice(0, 3);
      if (!kws.length) return;
      const sameName = norm(target.name).trim();
      cats = cats.map((c) =>
        norm(c.name).trim() === sameName && c.type === target.type
          ? { ...c, keywords: [...new Set([...(c.keywords || []), ...kws])] }
          : c
      );
    });
    const txs = valid.map((d) => ({
      id: uid(), type: d.type, amount: d.amount, description: d.description,
      categoryId: d.categoryId || null, accountId: d.accountId, date: d.date,
    }));
    onSave(txs, cats);
  }

  /* ---------- pantalla: subir ---------- */
  if (phase === "upload") {
    return (
      <div className="cc-overlay" onClick={onClose}>
        <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
          <div className="cc-grip" />
          <div className="cc-sheet-top">
            <h2>Importar screenshot</h2>
            <button className="cc-sheet-close" onClick={onClose}>×</button>
          </div>
          <p style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 16 }}>
            Sube screenshots de tu app bancaria o estado de cuenta. La IA extrae los movimientos y tú los revisas.
          </p>

          {config.accounts.length > 1 && (
            <div style={{ marginBottom: 14 }}>
              <label className="cc-label">Cuenta destino</label>
              <select className="cc-select" value={defaultAccountId} onChange={(e) => setDefaultAccountId(e.target.value)}
                style={{ borderColor: !defaultAccountId ? "var(--gold)" : "var(--line)" }}>
                <option value="">Elegir cuenta…</option>
                {config.accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
            </div>
          )}

          <input ref={inputRef} type="file" accept="image/*" multiple style={{ display: "none" }}
            onChange={(e) => addFiles(e.target.files)} />

          <button className="cc-card" onClick={() => inputRef.current?.click()}
            style={{ width: "100%", padding: 26, textAlign: "center", border: "2px dashed var(--line)",
              background: "var(--surface-2)", cursor: "pointer", marginBottom: 14 }}>
            <div style={{ fontSize: 30, marginBottom: 6 }}>📷</div>
            <div style={{ fontWeight: 600, fontSize: 14 }}>Toca para elegir imágenes</div>
            <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 4 }}>
              También puedes seleccionar varias a la vez
            </div>
          </button>

          {files.length > 0 && (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 8, marginBottom: 16 }}>
              {files.map((f, i) => (
                <div key={i} style={{ position: "relative", borderRadius: 10, overflow: "hidden", aspectRatio: "1", border: "1px solid var(--line)" }}>
                  <img src={f.previewUrl} alt="" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
                  <button onClick={() => removeFile(i)}
                    style={{ position: "absolute", top: 4, right: 4, width: 24, height: 24, borderRadius: "50%",
                      border: "none", background: "rgba(0,0,0,.7)", color: "#fff", cursor: "pointer", fontSize: 14, lineHeight: 1 }}>✕</button>
                </div>
              ))}
            </div>
          )}

          {error && (
            <div style={{ padding: 11, borderRadius: 10, background: "var(--coral-soft)", color: "var(--coral)", fontSize: 13, marginBottom: 14, fontWeight: 600 }}>
              {error}
            </div>
          )}

          <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 14, fontSize: 15 }}
            disabled={!files.length || !defaultAccountId} onClick={process}>
            ✨ Extraer movimientos
          </button>
        </div>
      </div>
    );
  }

  /* ---------- pantalla: procesando ---------- */
  if (phase === "processing") {
    return (
      <div className="cc-overlay">
        <div className="cc-sheet" style={{ padding: "40px 18px" }}>
          <div className="cc-grip" />
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <div style={{ fontSize: 36, marginBottom: 14 }}>🔍</div>
            <div className="cc-serif" style={{ fontSize: 20, fontWeight: 600, marginBottom: 6 }}>
              Leyendo tus movimientos…
            </div>
            <div style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 18 }}>
              La IA está revisando {files.length} imagen{files.length > 1 ? "es" : ""}
            </div>
            <div className="cc-dots"><span /><span /><span /></div>
          </div>
        </div>
      </div>
    );
  }

  /* ---------- pantalla: revisar ---------- */
  const okCount = drafts.filter((d) => d.selected && d.amount > 0).length;
  const accCats = (type) => config.categories.filter((c) => c.accountId === defaultAccountId && c.type === type);

  return (
    <ReviewScreen
      drafts={drafts}
      updateDraft={updateDraft}
      accCats={accCats}
      onBack={() => { setDrafts([]); setPhase("upload"); }}
      onSave={saveAll}
      onClose={onClose}
      sourceLabel="Encontré"
    />
  );
}

/* ============================ ESTADÍSTICAS =============================== */
/* ===== Gráfica de categorías versátil: pastel / dona / barras ===== */
const CHART_PALETTE = ["#5B6EE8", "#7C8BF5", "#60A5FA", "#5EEAD4", "#A78BFA", "#F0A868", "#E8849B", "#7E8AA0", "#86B98E", "#C9A24B"];

function CategoryChart({ rows, type, onPick }) {
  // rows: [{cat, amt}]
  const [chartType, setChartType] = useState(type === "income" ? "donut" : "bars");
  const [activeIdx, setActiveIdx] = useState(null);
  const total = rows.reduce((s, r) => s + r.amt, 0);
  if (!rows.length) return <div style={{ color: "var(--ink-soft)", fontSize: 13, padding: "8px 0 14px" }}>No hay datos en el periodo.</div>;

  const data = rows.map((r, i) => ({ ...r, color: CHART_PALETTE[i % CHART_PALETTE.length] }));
  const accentColor = type === "income" ? "var(--green)" : "var(--coral)";

  const TYPES = [["bars", "Barras"], ["pie", "Pastel"], ["donut", "Dona"]];

  return (
    <div>
      {/* switch de tipo */}
      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        {TYPES.map(([k, l]) => (
          <button key={k} onClick={() => { setChartType(k); setActiveIdx(null); }}
            style={{ padding: "5px 12px", borderRadius: 10, cursor: "pointer", fontFamily: "inherit",
              fontSize: 11.5, fontWeight: chartType === k ? 700 : 500,
              background: chartType === k ? "var(--ink)" : "var(--surface)",
              color: chartType === k ? "#fff" : "var(--ink-soft)",
              border: `1px solid ${chartType === k ? "var(--ink)" : "var(--line)"}` }}>
            {l}
          </button>
        ))}
      </div>

      {(chartType === "pie" || chartType === "donut") && (
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 14 }}>
          <CatPie data={data} total={total} donut={chartType === "donut"}
            active={activeIdx} onHover={setActiveIdx} />
          <div style={{ display: "flex", flexDirection: "column", gap: 2, width: "100%" }}>
            {data.map((d, i) => {
              const pct = total ? Math.round((d.amt / total) * 100) : 0;
              return (
                <button key={d.cat.id} onClick={() => onPick && onPick(d.cat.id)}
                  onMouseEnter={() => setActiveIdx(i)} onMouseLeave={() => setActiveIdx(null)}
                  style={{ display: "flex", alignItems: "center", gap: 9, padding: "8px 6px",
                    background: activeIdx === i ? "var(--surface)" : "transparent", border: "none",
                    borderRadius: 8, cursor: "pointer", fontFamily: "inherit", textAlign: "left", width: "100%",
                    transition: "background .12s" }}>
                  <span style={{ width: 10, height: 10, borderRadius: 3, background: d.color, flexShrink: 0 }} />
                  <span className="cc-emoji" style={{ fontSize: 15 }}>{d.cat.emoji}</span>
                  <span style={{ flex: 1, fontSize: 13, fontWeight: 600, color: "var(--ink)" }}>{d.cat.name}</span>
                  <span className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 13, color: accentColor }}>
                    {fmtBare(d.amt)}
                  </span>
                  <span style={{ fontSize: 10.5, color: "var(--ink-faint)", width: 30, textAlign: "right" }}>{pct}%</span>
                </button>
              );
            })}
          </div>
        </div>
      )}

      {chartType === "bars" && (
        <div style={{ display: "flex", flexDirection: "column" }}>
          {data.map((d) => {
            const pct = total ? Math.round((d.amt / total) * 100) : 0;
            const maxAmt = data[0].amt;
            const w = maxAmt ? (d.amt / maxAmt) * 100 : 0;
            return (
              <button key={d.cat.id} onClick={() => onPick && onPick(d.cat.id)}
                style={{ display: "flex", alignItems: "center", gap: 10, padding: "11px 0",
                  background: "transparent", border: "none", borderBottom: "1px solid var(--line-soft)",
                  cursor: "pointer", fontFamily: "inherit", textAlign: "left", width: "100%" }}>
                <div style={{ width: 34, height: 34, borderRadius: 10, background: "var(--surface)",
                  display: "flex", alignItems: "center", justifyContent: "center", fontSize: 17, flexShrink: 0 }}>
                  <span className="cc-emoji">{d.cat.emoji}</span>
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 600, fontSize: 13, color: "var(--ink)", letterSpacing: "-.01em" }}>{d.cat.name}</div>
                  <div style={{ height: 5, background: "var(--surface-2)", borderRadius: 99, overflow: "hidden", marginTop: 5 }}>
                    <div style={{ height: "100%", width: `${w}%`, background: type === "income" ? "var(--green)" : "var(--bar-fill)", borderRadius: 99 }} />
                  </div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div className="cc-num" style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 14, color: accentColor }}>
                    {fmtBare(d.amt)}<span style={{ fontSize: 10, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
                  </div>
                  <div style={{ fontSize: 10.5, color: "var(--ink-faint)", marginTop: 1 }}>{pct}%</div>
                </div>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* Pastel/dona interactivo con segmento activo resaltado */
function CatPie({ data, total, donut, active, onHover }) {
  const size = 180, cx = size / 2, cy = size / 2, r = 78, ir = donut ? 46 : 0;
  let acc = 0;
  const slices = data.map((d) => {
    const frac = total ? d.amt / total : 0;
    const start = acc; acc += frac;
    return { ...d, start, end: acc };
  });
  function arc(s, e, rad, innerRad) {
    const a0 = s * Math.PI * 2 - Math.PI / 2;
    const a1 = e * Math.PI * 2 - Math.PI / 2;
    const large = e - s > 0.5 ? 1 : 0;
    const x0 = cx + rad * Math.cos(a0), y0 = cy + rad * Math.sin(a0);
    const x1 = cx + rad * Math.cos(a1), y1 = cy + rad * Math.sin(a1);
    if (innerRad === 0) {
      return `M${cx},${cy} L${x0},${y0} A${rad},${rad} 0 ${large} 1 ${x1},${y1} Z`;
    }
    const xi0 = cx + innerRad * Math.cos(a0), yi0 = cy + innerRad * Math.sin(a0);
    const xi1 = cx + innerRad * Math.cos(a1), yi1 = cy + innerRad * Math.sin(a1);
    return `M${x0},${y0} A${rad},${rad} 0 ${large} 1 ${x1},${y1} L${xi1},${yi1} A${innerRad},${innerRad} 0 ${large} 0 ${xi0},${yi0} Z`;
  }
  const activeSlice = active != null ? slices[active] : null;
  return (
    <svg viewBox={`0 0 ${size} ${size}`} style={{ width: 180, height: 180, flexShrink: 0 }}>
      {slices.map((s, i) =>
        s.end - s.start < 0.001 ? null :
        <path key={i} d={arc(s.start, s.end, active === i ? r + 4 : r, ir)} fill={s.color}
          opacity={active == null || active === i ? 1 : 0.45}
          style={{ transition: "opacity .15s, d .15s", cursor: "pointer" }}
          onMouseEnter={() => onHover && onHover(i)} onMouseLeave={() => onHover && onHover(null)} />
      )}
      {donut && (
        <>
          <text x={cx} y={cy - 4} textAnchor="middle" fontFamily="Fraunces, serif" fontSize="12" fill="var(--ink-soft)">
            {activeSlice ? activeSlice.cat.name : "Total"}
          </text>
          <text x={cx} y={cy + 15} textAnchor="middle" fontFamily="Fraunces, serif" fontSize="16" fontWeight="600" fill="var(--ink)">
            {fmtBare(activeSlice ? activeSlice.amt : total)}
          </text>
        </>
      )}
    </svg>
  );
}

function Estadisticas({ config, txs, dateRange, onEdit, saveConfig }) {
  const [view, setView] = useState("all"); // all | accountId
  const [detail, setDetail] = useState(null); // {kind, label, color, txs, total}
  const [configOpen, setConfigOpen] = useState(false);

  // secciones configurables (orden + visibilidad)
  const STATS_DEFAULT = [
    { id: "summary", label: "Resumen ingresos/gastos", on: true },
    { id: "kpis", label: "Indicadores (KPIs)", on: true },
    { id: "incCats", label: "Ingresos por categoría", on: true },
    { id: "expCats", label: "Gastos por categoría", on: true },
    { id: "trend", label: "Evolución de saldo", on: true },
    { id: "topCat", label: "En lo que más gastaste", on: true },
  ];
  const statsSections = (() => {
    const saved = config.statsSections || [];
    // merge: respeta orden guardado, agrega nuevas que falten
    const byId = Object.fromEntries(saved.map((s) => [s.id, s]));
    const merged = saved
      .filter((s) => STATS_DEFAULT.some((d) => d.id === s.id))
      .map((s) => ({ ...STATS_DEFAULT.find((d) => d.id === s.id), ...s }));
    STATS_DEFAULT.forEach((d) => { if (!byId[d.id]) merged.push(d); });
    return merged.length ? merged : STATS_DEFAULT;
  })();
  const saveStatsSections = (next) => saveConfig({ ...config, statsSections: next });

  const scopedTxs = view === "all" ? txs : txs.filter((t) => t.accountId === view);
  const scopedInitial = view === "all"
    ? config.accounts.reduce((s, a) => s + (a.initialBalance || 0), 0)
    : (config.accounts.find((a) => a.id === view)?.initialBalance || 0);

  // ============== datos del rango ==============
  const rangeTxs = txsInRange(scopedTxs, dateRange);
  const rangeStat = statTxs(rangeTxs).all;
  const rangeIncome = rangeStat.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const rangeExpense = rangeStat.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  const rangeFlow = rangeIncome - rangeExpense;

  // categorías de gasto e ingreso por separado
  const expByCat = {};
  const incByCat = {};
  rangeStat.forEach((t) => {
    if (!t.categoryId) return;
    if (t.type === "expense") expByCat[t.categoryId] = (expByCat[t.categoryId] || 0) + t.amount;
    else incByCat[t.categoryId] = (incByCat[t.categoryId] || 0) + t.amount;
  });
  const expRows = Object.entries(expByCat)
    .map(([id, amt]) => ({ cat: config.categories.find((c) => c.id === id), amt }))
    .filter((x) => x.cat).sort((a, b) => b.amt - a.amt);
  const incRows = Object.entries(incByCat)
    .map(([id, amt]) => ({ cat: config.categories.find((c) => c.id === id), amt }))
    .filter((x) => x.cat).sort((a, b) => b.amt - a.amt);

  // pie de gastos (siempre del rango)
  const pieTotal = rangeExpense;

  // ============== evolución de saldo (90 días o todo el rango) ==============
  const { from: rfrom, to: rto } = resolveRange(dateRange);
  const sorted = [...scopedTxs].sort((a, b) => a.date.localeCompare(b.date));
  const balancePoints = [];
  let running = scopedInitial;
  for (const t of sorted) {
    if (t.date < rfrom) running += t.type === "income" ? t.amount : -t.amount;
  }
  const dayMap = new Map();
  dayMap.set(rfrom, running);
  for (const t of sorted) {
    if (t.date < rfrom || t.date > rto) continue;
    running += t.type === "income" ? t.amount : -t.amount;
    dayMap.set(t.date, running);
  }
  let lastVal = scopedInitial + sorted.filter((t) => t.date < rfrom).reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0);
  const startD = new Date(rfrom + "T12:00:00");
  const endD = new Date(rto + "T12:00:00");
  for (let d = new Date(startD); d <= endD; d.setDate(d.getDate() + 1)) {
    const k = d.toISOString().slice(0, 10);
    if (dayMap.has(k)) lastVal = dayMap.get(k);
    balancePoints.push({ date: k, val: lastVal });
  }

  // ============== KPIs ==============
  const days = Math.max(1, Math.round((endD - startD) / 86400000) + 1);
  const avgDaily = rangeExpense / days;
  const topCat = expRows[0];
  const txCount = rangeTxs.length;

  const showName = view === "all" ? "General" : (config.accounts.find((a) => a.id === view)?.name || "");

  // helpers de drill-down
  function openDetail(kind) {
    if (kind === "income") {
      const list = rangeStat.filter((t) => t.type === "income");
      setDetail({ kind: "income", label: "Todos los ingresos", color: "var(--green)",
        emoji: "📥", txs: list, total: rangeIncome });
    } else if (kind === "expense") {
      const list = rangeStat.filter((t) => t.type === "expense");
      setDetail({ kind: "expense", label: "Todos los gastos", color: "var(--coral)",
        emoji: "📤", txs: list, total: rangeExpense });
    }
  }
  function openCategoryDetail(catId) {
    const c = config.categories.find((x) => x.id === catId);
    if (!c) return;
    const list = rangeStat.filter((t) => t.categoryId === catId);
    const total = list.reduce((s, t) => s + t.amount, 0);
    setDetail({ kind: "category", label: c.name, color: c.type === "income" ? "var(--green)" : "var(--coral)",
      emoji: c.emoji, txs: list, total });
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      {/* selector de cuenta */}
      <div className="cc-fade">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div className="cc-label" style={{ marginBottom: 0 }}>Ver estadísticas de</div>
          <button className="cc-gear" onClick={() => setConfigOpen(true)} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, padding: "6px 12px", width: "auto" }}>
            ⚙️ Personalizar
          </button>
        </div>
        <div className="cc-scroll-x">
          <button className={`cc-acc-card ${view === "all" ? "on" : ""}`} onClick={() => setView("all")}
            style={{ minWidth: 120 }}>
            <div className="cc-acc-label">Todas</div>
            <div className="cc-acc-name">General</div>
          </button>
          {config.accounts.map((a) => (
            <button key={a.id} className={`cc-acc-card ${view === a.id ? "on" : ""}`} onClick={() => setView(a.id)}
              style={{ minWidth: 120 }}>
              <div className="cc-acc-label">🏦 Cuenta</div>
              <div className="cc-acc-name">{a.name}</div>
            </button>
          ))}
        </div>
      </div>

      {rangeTxs.length === 0 ? (
        <div className="cc-card" style={{ padding: 26, textAlign: "center" }}>
          <div style={{ fontSize: 30, marginBottom: 8 }}>📊</div>
          <div style={{ fontWeight: 600, fontSize: 15 }}>Sin movimientos en {showName}</div>
          <div style={{ fontSize: 13, color: "var(--ink-soft)", marginTop: 6 }}>
            Para el periodo: <b>{rangeLabel(dateRange)}</b>. Cambia el rango arriba o registra movimientos.
          </div>
        </div>
      ) : (
        <>
          {statsSections.filter((s) => s.on).map((s) => {
            if (s.id === "summary") return (
              <div key={s.id} className="cc-card" style={{ padding: 18 }}>
                <div className="cc-label" style={{ marginBottom: 12 }}>Resumen · {rangeLabel(dateRange)}</div>
                <div style={{ display: "flex", gap: 10 }}>
                  <button onClick={() => openDetail("income")}
                    style={{ flex: 1, padding: "14px 15px", background: "var(--surface)",
                      border: "1px solid var(--line-soft)", borderRadius: 14, cursor: "pointer", textAlign: "left",
                      fontFamily: "inherit", display: "flex", flexDirection: "column", gap: 4 }}>
                    <div style={{ fontSize: 10.5, fontWeight: 600, color: "var(--ink-faint)",
                      textTransform: "uppercase", letterSpacing: ".06em" }}>Ingresos</div>
                    <div className="cc-serif cc-num" style={{ fontSize: 22, fontWeight: 500, color: "var(--green)" }}>{fmtBare(rangeIncome)}</div>
                    <div style={{ fontSize: 11, color: "var(--ink-faint)" }}>Tocar para detalle ▸</div>
                  </button>
                  <button onClick={() => openDetail("expense")}
                    style={{ flex: 1, padding: "14px 15px", background: "var(--surface)",
                      border: "1px solid var(--line-soft)", borderRadius: 14, cursor: "pointer", textAlign: "left",
                      fontFamily: "inherit", display: "flex", flexDirection: "column", gap: 4 }}>
                    <div style={{ fontSize: 10.5, fontWeight: 600, color: "var(--ink-faint)",
                      textTransform: "uppercase", letterSpacing: ".06em" }}>Gastos</div>
                    <div className="cc-serif cc-num" style={{ fontSize: 22, fontWeight: 500, color: "var(--coral)" }}>{fmtBare(rangeExpense)}</div>
                    <div style={{ fontSize: 11, color: "var(--ink-faint)" }}>Tocar para detalle ▸</div>
                  </button>
                </div>
                <div style={{ marginTop: 12, paddingTop: 12, borderTop: "1px solid var(--line-soft)",
                  display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 13, fontWeight: 600, color: "var(--ink-soft)" }}>Flujo neto</span>
                  <span className="cc-serif cc-num" style={{ fontSize: 18, fontWeight: 500,
                    color: rangeFlow >= 0 ? "var(--ink)" : "var(--coral)" }}>
                    {rangeFlow >= 0 ? "+" : "−"}{fmtBare(Math.abs(rangeFlow))}<span style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 11, fontWeight: 300, color: "var(--ink-faint)", marginLeft: 3 }}>mxn</span>
                  </span>
                </div>
              </div>
            );

            if (s.id === "kpis") return (
              <div key={s.id} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <KpiCard label="Gasto promedio diario" value={fmt(avgDaily)} color="var(--coral)" />
                <KpiCard label="Movimientos en el periodo" value={String(txCount)} color="var(--ink)" />
              </div>
            );

            if (s.id === "incCats" && incRows.length > 0) return (
              <div key={s.id} className="cc-card" style={{ padding: 18 }}>
                <div className="cc-label" style={{ marginBottom: 12 }}>Ingresos por categoría</div>
                <CategoryChart rows={incRows} type="income" onPick={openCategoryDetail} />
              </div>
            );

            if (s.id === "expCats" && expRows.length > 0) return (
              <div key={s.id} className="cc-card" style={{ padding: 18 }}>
                <div className="cc-label" style={{ marginBottom: 12 }}>Gastos por categoría</div>
                <CategoryChart rows={expRows} type="expense" onPick={openCategoryDetail} />
              </div>
            );

            if (s.id === "trend" && balancePoints.length >= 2) return (
              <div key={s.id} className="cc-card" style={{ padding: 18 }}>
                <div className="cc-label" style={{ marginBottom: 10 }}>Saldo · {rangeLabel(dateRange)}</div>
                <LineChart points={balancePoints} />
                <div style={{ fontSize: 11, color: "var(--ink-faint)", marginTop: 8, textAlign: "center" }}>
                  Desliza sobre la gráfica para ver cualquier día
                </div>
              </div>
            );

            if (s.id === "topCat" && topCat) return (
              <div key={s.id} className="cc-card" style={{ padding: 18 }}>
                <div className="cc-label" style={{ marginBottom: 6 }}>En lo que más gastaste</div>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  <span className="cc-emoji" style={{ fontSize: 32 }}>{topCat.cat.emoji}</span>
                  <div style={{ flex: 1 }}>
                    <div className="cc-serif" style={{ fontSize: 19, fontWeight: 600 }}>{topCat.cat.name}</div>
                    <div className="cc-num" style={{ fontSize: 14, color: "var(--ink-soft)" }}>
                      {fmt(topCat.amt)} · {pieTotal ? Math.round((topCat.amt / pieTotal) * 100) : 0}% de tus gastos
                    </div>
                  </div>
                </div>
              </div>
            );

            return null;
          })}
        </>
      )}

      {configOpen && (
        <StatsConfigModal
          sections={statsSections}
          onClose={() => setConfigOpen(false)}
          onSave={(next) => { saveStatsSections(next); setConfigOpen(false); }}
        />
      )}

      {detail && (
        <DetailModal config={config} detail={detail} dateRange={dateRange}
          onClose={() => setDetail(null)}
          onEditTx={(t) => { setDetail(null); onEdit(t); }} />
      )}
    </div>
  );
}

/* ===== Modal: personalizar secciones de Estadísticas (reordenar + on/off) ===== */
function StatsConfigModal({ sections, onClose, onSave }) {
  const [items, setItems] = useState(sections.map((s) => ({ ...s })));

  const move = (i, dir) => {
    const j = i + dir;
    if (j < 0 || j >= items.length) return;
    const next = [...items];
    [next[i], next[j]] = [next[j], next[i]];
    setItems(next);
  };
  const toggle = (i) => {
    setItems(items.map((s, idx) => (idx === i ? { ...s, on: !s.on } : s)));
  };

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>{t("customizeStatsTitle")}</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 16 }}>
          Reordena con las flechas y muestra u oculta cada sección.
        </p>

        <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 18 }}>
          {items.map((s, i) => (
            <div key={s.id} className="cc-card" style={{ padding: "11px 14px", display: "flex", alignItems: "center", gap: 10, opacity: s.on ? 1 : 0.55 }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                <button onClick={() => move(i, -1)} disabled={i === 0}
                  style={{ background: "var(--surface)", border: "1px solid var(--line)", borderRadius: 6, width: 26, height: 20, cursor: i === 0 ? "default" : "pointer", opacity: i === 0 ? 0.3 : 1, fontSize: 11, lineHeight: 1, color: "var(--ink)" }}>▲</button>
                <button onClick={() => move(i, 1)} disabled={i === items.length - 1}
                  style={{ background: "var(--surface)", border: "1px solid var(--line)", borderRadius: 6, width: 26, height: 20, cursor: i === items.length - 1 ? "default" : "pointer", opacity: i === items.length - 1 ? 0.3 : 1, fontSize: 11, lineHeight: 1, color: "var(--ink)" }}>▼</button>
              </div>
              <span style={{ flex: 1, fontSize: 14, fontWeight: 600, color: "var(--ink)" }}>{s.label}</span>
              <button onClick={() => toggle(i)}
                style={{ width: 46, height: 27, borderRadius: 99, border: "none", cursor: "pointer", position: "relative",
                  background: s.on ? "var(--green)" : "var(--surface-2)", transition: "background .15s" }}>
                <span style={{ position: "absolute", top: 3, left: s.on ? 22 : 3, width: 21, height: 21, borderRadius: "50%",
                  background: "#fff", transition: "left .15s", boxShadow: "0 1px 3px rgba(0,0,0,.2)" }} />
              </button>
            </div>
          ))}
        </div>

        <button className="cc-btn cc-btn-primary" style={{ width: "100%", padding: 14 }}
          onClick={() => onSave(items)}>Guardar</button>
      </div>
    </div>
  );
}

function KpiCard({ label, value, sub, color }) {
  return (
    <div className="cc-card cc-card-boxed" style={{ padding: 14, paddingTop: 14 }}>
      <div className="cc-label" style={{ marginBottom: 6 }}>{label}</div>
      <div className="cc-serif cc-num" style={{ fontSize: 22, fontWeight: 500, color: color || "var(--ink)", letterSpacing: "-.025em", lineHeight: 1 }}>{value}</div>
      {sub && <div className="cc-num" style={{ fontSize: 11, color: "var(--ink-soft)", marginTop: 4 }}>{sub}</div>}
    </div>
  );
}

/* ----------------------- gráfica de línea (SVG) -------------------------- */
function LineChart({ points, area: showArea = true }) {
  const [hover, setHover] = useState(null); // index hovered
  if (!points || points.length < 2) {
    return <div style={{ fontSize: 13, color: "var(--ink-soft)", padding: "20px 0" }}>Datos insuficientes.</div>;
  }
  const W = 600, H = 200, P = 10, PB = 22; // PB: padding inferior para etiquetas
  const vals = points.map((p) => p.val);
  const min = Math.min(...vals);
  const max = Math.max(...vals);
  const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
  const range = max - min || 1;
  const xOf = (i) => P + (i / (points.length - 1)) * (W - P * 2);
  const yOf = (v) => H - PB - ((v - min) / range) * (H - P - PB);
  const path = points.map((p, i) => `${i === 0 ? "M" : "L"}${xOf(i).toFixed(1)},${yOf(p.val).toFixed(1)}`).join(" ");
  const areaPath = `${path} L${xOf(points.length - 1).toFixed(1)},${H - PB} L${P},${H - PB} Z`;
  const last = points[points.length - 1].val;
  const first = points[0].val;
  const up = last >= first;
  const stroke = up ? "var(--green)" : "var(--coral)";
  const avgY = yOf(avg);

  const maxIdx = vals.indexOf(max);
  const minIdx = vals.indexOf(min);

  const handleMove = (e) => {
    const svg = e.currentTarget;
    const rect = svg.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const x = ((clientX - rect.left) / rect.width) * W;
    const i = Math.round(((x - P) / (W - P * 2)) * (points.length - 1));
    setHover(Math.max(0, Math.min(points.length - 1, i)));
  };

  const hp = hover != null ? points[hover] : null;

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8, fontSize: 12, color: "var(--ink-soft)" }}>
        <span>{hp ? hp.date : points[0].date}</span>
        <span className="cc-num" style={{ fontWeight: 700, color: stroke, fontSize: 14 }}>
          {hp ? fmt(hp.val) : fmt(last)}
        </span>
        <span>{hp ? "" : "hoy"}</span>
      </div>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block", touchAction: "none" }}
        onMouseMove={handleMove} onMouseLeave={() => setHover(null)}
        onTouchStart={handleMove} onTouchMove={handleMove} onTouchEnd={() => setHover(null)}>
        <defs>
          <linearGradient id="ccLine" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor={stroke} stopOpacity="0.22" />
            <stop offset="100%" stopColor={stroke} stopOpacity="0" />
          </linearGradient>
        </defs>
        {/* promedio */}
        <line x1={P} y1={avgY} x2={W - P} y2={avgY} stroke="var(--ink-faint)" strokeWidth="1" strokeDasharray="4 4" opacity="0.6" />
        <text x={W - P} y={avgY - 4} textAnchor="end" fontSize="10" fill="var(--ink-faint)">prom {fmtBare(avg)}</text>
        {showArea && <path d={areaPath} fill="url(#ccLine)" />}
        <path d={path} fill="none" stroke={stroke} strokeWidth="2.2" strokeLinejoin="round" strokeLinecap="round" />
        {/* máximo y mínimo */}
        <circle cx={xOf(maxIdx)} cy={yOf(max)} r="3.5" fill="var(--green)" />
        <text x={xOf(maxIdx)} y={yOf(max) - 7} textAnchor="middle" fontSize="10" fontWeight="600" fill="var(--green)">{fmtBare(max)}</text>
        <circle cx={xOf(minIdx)} cy={yOf(min)} r="3.5" fill="var(--coral)" />
        <text x={xOf(minIdx)} y={yOf(min) + 14} textAnchor="middle" fontSize="10" fontWeight="600" fill="var(--coral)">{fmtBare(min)}</text>
        {/* cursor hover */}
        {hp && (
          <>
            <line x1={xOf(hover)} y1={P} x2={xOf(hover)} y2={H - PB} stroke="var(--ink-soft)" strokeWidth="1" opacity="0.4" />
            <circle cx={xOf(hover)} cy={yOf(hp.val)} r="5" fill="#fff" stroke={stroke} strokeWidth="2.5" />
          </>
        )}
      </svg>
    </div>
  );
}

/* --------------------- gráfica de pastel (SVG donut) --------------------- */
const PIE_COLORS = ["#B8482A", "#B0863A", "#2E6F4E", "#7B5E2E", "#9E3F26", "#5C7A4C", "#A07344", "#7A4630"];

function PieChart({ data, total }) {
  const size = 160, cx = size / 2, cy = size / 2, r = 70, ir = 42;
  let acc = 0;
  const slices = data.map((d, i) => {
    const frac = total ? d.amt / total : 0;
    const start = acc; acc += frac;
    return { ...d, start, end: acc, color: PIE_COLORS[i % PIE_COLORS.length] };
  });

  function arc(s, e) {
    const a0 = s * Math.PI * 2 - Math.PI / 2;
    const a1 = e * Math.PI * 2 - Math.PI / 2;
    const large = e - s > 0.5 ? 1 : 0;
    const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
    const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
    const xi0 = cx + ir * Math.cos(a0), yi0 = cy + ir * Math.sin(a0);
    const xi1 = cx + ir * Math.cos(a1), yi1 = cy + ir * Math.sin(a1);
    return `M${x0},${y0} A${r},${r} 0 ${large} 1 ${x1},${y1} L${xi1},${yi1} A${ir},${ir} 0 ${large} 0 ${xi0},${yi0} Z`;
  }

  return (
    <svg viewBox={`0 0 ${size} ${size}`} style={{ width: 160, height: 160, flexShrink: 0 }}>
      {slices.map((s, i) =>
        s.end - s.start < 0.001 ? null :
        <path key={i} d={arc(s.start, s.end)} fill={s.color} />
      )}
      <text x={cx} y={cy - 4} textAnchor="middle" fontFamily="Fraunces, serif" fontSize="13" fill="var(--ink-soft)">Total</text>
      <text x={cx} y={cy + 14} textAnchor="middle" fontFamily="Fraunces, serif" fontSize="15" fontWeight="600" fill="var(--ink)">{fmt(total)}</text>
    </svg>
  );
}

/* ---------------------- gráfica de barras (SVG) -------------------------- */
function BarsChart({ bars }) {
  const W = 600, H = 180, P = 18;
  const max = Math.max(1, ...bars.map((b) => Math.max(b.expense, b.income)));
  const groupW = (W - P * 2) / bars.length;
  const barW = groupW * 0.38;
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }}>
      {bars.map((b, i) => {
        const cx = P + groupW * i + groupW / 2;
        const hi = (b.income / max) * (H - P * 2);
        const he = (b.expense / max) * (H - P * 2);
        return (
          <g key={i}>
            <rect x={cx - barW - 2} y={H - P - hi} width={barW} height={hi} fill="var(--green)" rx="3" />
            <rect x={cx + 2} y={H - P - he} width={barW} height={he} fill="var(--coral)" rx="3" />
            <text x={cx} y={H - 4} textAnchor="middle" fontSize="10" fill="var(--ink-soft)" fontFamily="Montserrat, sans-serif">
              {monthLabel(b.key).split(" ")[0]}
            </text>
          </g>
        );
      })}
    </svg>
  );
}

/* ============= MiniChart para el chat asistente ========================== */
/* Recibe un payload tipo:
   { kind: "line", title, xLabels:["sem 1",...], series:[{name,color,values:[]}], yLabel? }
   { kind: "bars", title, xLabels, series:[{name,color,values}] }
   { kind: "pie",  title, segments:[{label, value, color?}] }
*/
function MiniChart({ data }) {
  if (!data || !data.kind) return null;
  const title = data.title || "";
  return (
    <div style={{ background: "var(--surface)", border: "1px solid var(--line)", borderRadius: 12, padding: 12, marginTop: 6 }}>
      {title && <div style={{ fontSize: 12, fontWeight: 700, color: "var(--ink-soft)", textTransform: "uppercase", letterSpacing: ".06em", marginBottom: 8 }}>{title}</div>}
      {data.kind === "line" && <MiniLine series={data.series || []} xLabels={data.xLabels || []} />}
      {data.kind === "bars" && <MiniBars series={data.series || []} xLabels={data.xLabels || []} />}
      {data.kind === "pie" && <MiniPie segments={data.segments || []} />}
      {/* leyenda */}
      {(data.kind === "line" || data.kind === "bars") && data.series && data.series.length > 0 && (
        <div style={{ display: "flex", gap: 14, fontSize: 11.5, color: "var(--ink-soft)", marginTop: 8, flexWrap: "wrap" }}>
          {data.series.map((s, i) => (
            <span key={i}>
              <span style={{ display: "inline-block", width: 9, height: 9, background: s.color || PIE_COLORS[i % PIE_COLORS.length], borderRadius: 2, marginRight: 5, verticalAlign: "middle" }} />
              {s.name}
            </span>
          ))}
        </div>
      )}
    </div>
  );
}

/* helper para abreviar montos en ejes ($1.2k, $34k) */
function shortMoney(v) {
  const abs = Math.abs(v);
  const sign = v < 0 ? "-" : "";
  if (abs >= 1_000_000) return `${sign}$${(abs / 1_000_000).toFixed(abs >= 10_000_000 ? 0 : 1)}M`;
  if (abs >= 1_000) return `${sign}$${(abs / 1_000).toFixed(abs >= 10_000 ? 0 : 1)}k`;
  return `${sign}$${Math.round(abs)}`;
}

/* genera ticks "redondos" para el eje Y */
function niceTicks(min, max, count = 4) {
  if (min === max) { min -= 1; max += 1; }
  const range = max - min;
  const rough = range / count;
  const pow = Math.pow(10, Math.floor(Math.log10(rough)));
  const norm = rough / pow;
  let step;
  if (norm < 1.5) step = 1;
  else if (norm < 3) step = 2;
  else if (norm < 7) step = 5;
  else step = 10;
  step *= pow;
  const niceMin = Math.floor(min / step) * step;
  const niceMax = Math.ceil(max / step) * step;
  const ticks = [];
  for (let v = niceMin; v <= niceMax + 0.0001; v += step) ticks.push(v);
  return { ticks, min: niceMin, max: niceMax };
}

function MiniLine({ series, xLabels }) {
  const [hover, setHover] = useState(null); // índice de bucket bajo el cursor
  if (!series.length) return null;
  const W = 600, H = 220, PL = 50, PR = 14, PT = 14, PB = 38;
  const all = series.flatMap((s) => s.values || []);
  if (!all.length) return null;
  const rawMin = Math.min(0, ...all);
  const rawMax = Math.max(...all, 1);
  const { ticks, min, max } = niceTicks(rawMin, rawMax, 4);
  const range = max - min || 1;
  const n = xLabels.length;
  const innerW = W - PL - PR;
  const innerH = H - PT - PB;
  const xOf = (i) => PL + (n > 1 ? (i / (n - 1)) * innerW : innerW / 2);
  const yOf = (v) => PT + innerH - ((v - min) / range) * innerH;
  const colWidth = innerW / Math.max(n, 1);

  return (
    <div style={{ position: "relative", width: "100%" }}>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }}
        onMouseLeave={() => setHover(null)}>
        {/* grid horizontal */}
        {ticks.map((tv, i) => (
          <g key={i}>
            <line x1={PL} x2={W - PR} y1={yOf(tv)} y2={yOf(tv)}
              stroke={tv === 0 ? "var(--ink-soft)" : "var(--line)"}
              strokeOpacity={tv === 0 ? 0.5 : 1}
              strokeDasharray={tv === 0 ? "none" : "3 4"} />
            <text x={PL - 6} y={yOf(tv) + 3} textAnchor="end" fontSize="10"
              fill="var(--ink-soft)" fontFamily="Montserrat, sans-serif">{shortMoney(tv)}</text>
          </g>
        ))}
        {/* etiquetas eje X */}
        {xLabels.map((lbl, i) => (
          <text key={i} x={xOf(i)} y={H - 16} textAnchor="middle" fontSize="11"
            fill="var(--ink-soft)" fontFamily="Montserrat, sans-serif">{lbl}</text>
        ))}
        {/* línea guía vertical */}
        {hover != null && (
          <line x1={xOf(hover)} x2={xOf(hover)} y1={PT} y2={PT + innerH}
            stroke="var(--ink-soft)" strokeDasharray="3 3" strokeOpacity="0.5" />
        )}
        {/* series */}
        {series.map((s, si) => {
          const color = s.color || PIE_COLORS[si % PIE_COLORS.length];
          const vals = s.values || [];
          const pts = vals.map((v, i) => `${i === 0 ? "M" : "L"}${xOf(i).toFixed(1)},${yOf(v).toFixed(1)}`).join(" ");
          return (
            <g key={si}>
              <path d={pts} fill="none" stroke={color} strokeWidth="2.4" strokeLinejoin="round" strokeLinecap="round" />
              {vals.map((v, i) => (
                <g key={i}>
                  <circle cx={xOf(i)} cy={yOf(v)} r={hover === i ? 5 : 3.5} fill={color}
                    stroke="var(--surface)" strokeWidth="1.5" />
                </g>
              ))}
            </g>
          );
        })}
        {/* columnas invisibles para captar hover/touch */}
        {xLabels.map((_, i) => {
          const cx = xOf(i);
          return (
            <rect key={i} x={cx - colWidth / 2} y={PT} width={colWidth} height={innerH}
              fill="transparent" style={{ cursor: "pointer" }}
              onMouseEnter={() => setHover(i)}
              onTouchStart={() => setHover(i)} />
          );
        })}
      </svg>
      {/* tooltip */}
      {hover != null && (
        <ChartTooltip
          label={xLabels[hover]}
          rows={series.map((s, si) => ({
            name: s.name, color: s.color || PIE_COLORS[si % PIE_COLORS.length],
            value: (s.values || [])[hover] || 0,
          }))}
          x={(xOf(hover) / W) * 100}
        />
      )}
    </div>
  );
}

function MiniBars({ series, xLabels }) {
  const [hover, setHover] = useState(null);
  if (!series.length) return null;
  const W = 600, H = 220, PL = 50, PR = 14, PT = 14, PB = 38;
  const all = series.flatMap((s) => s.values || []);
  if (!all.length) return null;
  const { ticks, min, max } = niceTicks(0, Math.max(...all, 1), 4);
  const range = max - min || 1;
  const n = xLabels.length;
  const innerW = W - PL - PR;
  const innerH = H - PT - PB;
  const groupW = innerW / Math.max(n, 1);
  const sCount = series.length;
  const barW = Math.min(30, (groupW * 0.7) / sCount);
  const cxOf = (i) => PL + groupW * i + groupW / 2;
  const yOf = (v) => PT + innerH - ((v - min) / range) * innerH;

  return (
    <div style={{ position: "relative", width: "100%" }}>
      <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%", height: "auto", display: "block" }}
        onMouseLeave={() => setHover(null)}>
        {/* grid horizontal */}
        {ticks.map((tv, i) => (
          <g key={i}>
            <line x1={PL} x2={W - PR} y1={yOf(tv)} y2={yOf(tv)}
              stroke="var(--line)" strokeDasharray={tv === 0 ? "none" : "3 4"} />
            <text x={PL - 6} y={yOf(tv) + 3} textAnchor="end" fontSize="10"
              fill="var(--ink-soft)" fontFamily="Montserrat, sans-serif">{shortMoney(tv)}</text>
          </g>
        ))}
        {/* highlight columna activa */}
        {hover != null && (
          <rect x={cxOf(hover) - groupW / 2} y={PT} width={groupW} height={innerH}
            fill="var(--ink)" fillOpacity="0.04" />
        )}
        {/* barras */}
        {xLabels.map((lbl, i) => {
          const cx = cxOf(i);
          return (
            <g key={i}>
              {series.map((s, si) => {
                const color = s.color || PIE_COLORS[si % PIE_COLORS.length];
                const v = (s.values && s.values[i]) || 0;
                const h = (v / max) * innerH;
                const x = cx + (si - (sCount - 1) / 2) * barW - barW / 2;
                return <rect key={si} x={x} y={H - PB - h} width={barW - 2} height={Math.max(h, 0)}
                  fill={color} rx="3" opacity={hover != null && hover !== i ? 0.55 : 1} />;
              })}
              <text x={cx} y={H - 16} textAnchor="middle" fontSize="11"
                fill="var(--ink-soft)" fontFamily="Montserrat, sans-serif">{lbl}</text>
            </g>
          );
        })}
        {/* columnas invisibles para hover */}
        {xLabels.map((_, i) => (
          <rect key={i} x={cxOf(i) - groupW / 2} y={PT} width={groupW} height={innerH}
            fill="transparent" style={{ cursor: "pointer" }}
            onMouseEnter={() => setHover(i)}
            onTouchStart={() => setHover(i)} />
        ))}
      </svg>
      {hover != null && (
        <ChartTooltip
          label={xLabels[hover]}
          rows={series.map((s, si) => ({
            name: s.name, color: s.color || PIE_COLORS[si % PIE_COLORS.length],
            value: (s.values || [])[hover] || 0,
          }))}
          x={(cxOf(hover) / W) * 100}
        />
      )}
    </div>
  );
}

/* tooltip flotante posicionado en porcentaje de ancho del contenedor */
function ChartTooltip({ label, rows, x }) {
  // si está a la derecha, anclar a la derecha; si a la izquierda, anclar a la izquierda
  const anchorRight = x > 60;
  return (
    <div style={{
      position: "absolute",
      top: 6,
      [anchorRight ? "right" : "left"]: anchorRight ? `${100 - x + 2}%` : `${x + 2}%`,
      maxWidth: "44%",
      background: "var(--ink)", color: "var(--surface)",
      padding: "8px 11px", borderRadius: 9,
      fontSize: 12, lineHeight: 1.4,
      boxShadow: "0 6px 20px rgba(0,0,0,.25)",
      pointerEvents: "none",
      zIndex: 5,
      whiteSpace: "nowrap",
    }}>
      <div style={{ fontWeight: 700, marginBottom: 4, fontSize: 11, opacity: 0.75, textTransform: "uppercase", letterSpacing: ".05em" }}>{label}</div>
      {rows.map((r, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 7, marginTop: i === 0 ? 0 : 2 }}>
          <span style={{ width: 8, height: 8, borderRadius: 2, background: r.color, flexShrink: 0 }} />
          <span style={{ flex: 1, opacity: 0.8 }}>{r.name}</span>
          <span style={{ fontWeight: 700, fontVariantNumeric: "tabular-nums" }}>{fmt(r.value)}</span>
        </div>
      ))}
    </div>
  );
}

function MiniPie({ segments }) {
  const total = segments.reduce((s, x) => s + (x.value || 0), 0);
  if (total <= 0) return <div style={{ fontSize: 13, color: "var(--ink-soft)" }}>Sin datos.</div>;
  const size = 160, cx = size / 2, cy = size / 2, r = 70, ir = 42;
  let acc = 0;
  const slices = segments.map((seg, i) => {
    const frac = seg.value / total;
    const start = acc; acc += frac;
    return { ...seg, start, end: acc, color: seg.color || PIE_COLORS[i % PIE_COLORS.length] };
  });
  function arc(s, e) {
    const a0 = s * Math.PI * 2 - Math.PI / 2;
    const a1 = e * Math.PI * 2 - Math.PI / 2;
    const large = e - s > 0.5 ? 1 : 0;
    const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
    const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
    const xi0 = cx + ir * Math.cos(a0), yi0 = cy + ir * Math.sin(a0);
    const xi1 = cx + ir * Math.cos(a1), yi1 = cy + ir * Math.sin(a1);
    return `M${x0},${y0} A${r},${r} 0 ${large} 1 ${x1},${y1} L${xi1},${yi1} A${ir},${ir} 0 ${large} 0 ${xi0},${yi0} Z`;
  }
  return (
    <div style={{ display: "flex", gap: 14, alignItems: "center", flexWrap: "wrap" }}>
      <svg viewBox={`0 0 ${size} ${size}`} style={{ width: 140, height: 140, flexShrink: 0 }}>
        {slices.map((s, i) => s.end - s.start < 0.001 ? null : <path key={i} d={arc(s.start, s.end)} fill={s.color} />)}
        <text x={cx} y={cy + 5} textAnchor="middle" fontFamily="Fraunces, serif" fontSize="14" fontWeight="600" fill="var(--ink)">{fmt(total)}</text>
      </svg>
      <div style={{ flex: 1, minWidth: 140, display: "flex", flexDirection: "column", gap: 5 }}>
        {slices.slice(0, 6).map((s, i) => {
          const pct = Math.round((s.value / total) * 100);
          return (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12.5 }}>
              <span style={{ width: 10, height: 10, borderRadius: 2, background: s.color, flexShrink: 0 }} />
              <span style={{ flex: 1, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{s.label}</span>
              <span className="cc-num" style={{ fontWeight: 700 }}>{pct}%</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ===================== MODAL: IMPORTAR DESDE EXCEL ======================= */
function ExcelImportModal({ config, txs, onClose, onSave }) {
  const [file, setFile] = useState(null);
  const [rows, setRows] = useState([]); // todas las filas crudas
  const [defaultAccountId, setDefaultAccountId] = useState(
    config.accounts.length === 1 ? config.accounts[0].id : ""
  );
  const [phase, setPhase] = useState("upload"); // upload | processing | review
  const [drafts, setDrafts] = useState([]);
  const [error, setError] = useState(null);
  const inputRef = useRef(null);

  async function handleFile(f) {
    if (!f) return;
    setError(null);
    setFile(f);
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
      const filtered = json.filter((r) => Array.isArray(r) && r.some((c) => c !== "" && c != null));
      if (filtered.length < 2) {
        setError("El archivo está vacío o no tiene suficientes filas.");
        return;
      }
      setRows(filtered);
    } catch (e) {
      console.error(e);
      setError("No pude leer el archivo. ¿Es un Excel válido (.xlsx, .xls, .csv)?");
    }
  }

  async function process() {
    if (!rows.length || !defaultAccountId) return;
    setPhase("processing");
    setError(null);
    try {
      // mandar sample a Claude para detectar columnas
      const sample = rows.slice(0, Math.min(15, rows.length));
      const cats = config.categories
        .filter((c) => c.accountId === defaultAccountId)
        .map((c) => ({ name: c.name, type: c.type }));

      const sys = `Analiza una hoja de cálculo de movimientos bancarios y dime qué columna es cada cosa.
Te paso las primeras filas como arreglo de arreglos. La fila 0 puede ser encabezado (o no).

Devuelve UN SOLO objeto JSON, sin markdown:
{
  "headerRow": <índice de la fila de encabezado, o null si no hay>,
  "columns": {
    "date": <índice de columna de fecha, o null>,
    "description": <índice de columna de concepto/descripción, o null>,
    "amount": <índice de columna de monto, o null>,
    "type": <índice de columna que indica tipo (gasto/ingreso/cargo/abono), o null>,
    "expenseColumn": <índice si hay UNA columna SOLO para gastos (cargo/débito), o null>,
    "incomeColumn": <índice si hay UNA columna SOLO para ingresos (abono/crédito), o null>
  },
  "amountSign": "negative-is-expense" | "positive-is-expense" | "separate-columns" | "type-column",
  "dateFormat": "<descripción breve del formato observado, ej: DD/MM/YYYY>"
}

REGLAS:
- Si hay UNA sola columna de monto y los gastos son negativos → "negative-is-expense"
- Si hay UNA sola columna y todos son positivos pero una columna aparte dice tipo (gasto/ingreso) → "type-column"
- Si hay DOS columnas separadas (una para cargos y otra para abonos) → "separate-columns" + llena expenseColumn e incomeColumn
- Si no estás seguro, usa null en los campos.`;

      const userMsg = `Filas: ${JSON.stringify(sample)}`;
      const raw = await callClaude(sys, [{ role: "user", content: userMsg }]);
      const mapping = parseJSON(raw);

      // construir drafts
      const headerIdx = mapping.headerRow != null ? mapping.headerRow : -1;
      const dataRows = rows.filter((_, i) => i !== headerIdx);
      const sign = mapping.amountSign;
      const cols = mapping.columns || {};

      const ds = [];
      dataRows.forEach((r, idx) => {
        if (!Array.isArray(r)) return;
        const dateRaw = cols.date != null ? r[cols.date] : "";
        const desc = cols.description != null ? String(r[cols.description] || "") : "";

        let amount = 0, type = "expense";
        if (sign === "separate-columns" && cols.expenseColumn != null && cols.incomeColumn != null) {
          const e = parseAmount(r[cols.expenseColumn]);
          const inc = parseAmount(r[cols.incomeColumn]);
          if (inc > 0) { amount = inc; type = "income"; }
          else if (e > 0) { amount = e; type = "expense"; }
        } else if (sign === "type-column" && cols.amount != null && cols.type != null) {
          amount = Math.abs(parseAmount(r[cols.amount]));
          const tt = String(r[cols.type] || "").toLowerCase();
          type = /ingres|abono|cr[eé]dito|depo/.test(tt) ? "income" : "expense";
        } else if (cols.amount != null) {
          const a = parseAmount(r[cols.amount]);
          if (sign === "positive-is-expense") {
            amount = Math.abs(a); type = a >= 0 ? "expense" : "income";
          } else {
            amount = Math.abs(a); type = a < 0 ? "expense" : "income";
          }
        }
        if (amount <= 0) return; // sin monto, descartar

        const date = parseDate(dateRaw) || today();
        // auto-categorizar con keywords locales
        const nd = norm(desc);
        let catId = "";
        const accCats = config.categories.filter((c) => c.accountId === defaultAccountId && c.type === type);
        for (const c of accCats) {
          if ((c.keywords || []).some((kw) => nd.includes(norm(kw)))) { catId = c.id; break; }
        }
        const candidate = { date, amount, type, description: desc.trim(), accountId: defaultAccountId };
        const dup = findDuplicate(candidate, txs || []);
        ds.push({
          tempId: "x" + idx, ...candidate, categoryId: catId,
          selected: !dup,
          duplicate: dup ? { date: dup.date, amount: dup.amount, description: dup.description } : null,
        });
      });

      if (!ds.length) {
        setError("No detecté movimientos válidos. Revisa que el archivo tenga columnas de fecha y monto.");
        setPhase("upload");
        return;
      }
      setDrafts(ds);
      setPhase("review");
    } catch (e) {
      console.error(e);
      setError("Hubo un problema procesando el archivo. Inténtalo de nuevo.");
      setPhase("upload");
    }
  }

  function updateDraft(tempId, patch) {
    setDrafts((prev) => prev.map((d) => (d.tempId === tempId ? { ...d, ...patch } : d)));
  }

  function saveAll() {
    const valid = drafts.filter((d) => d.selected && d.amount > 0);
    if (!valid.length) return;
    let cats = config.categories;
    valid.forEach((d) => {
      if (!d.categoryId || !d.description) return;
      const target = cats.find((c) => c.id === d.categoryId);
      if (!target) return;
      const kws = extractKW(d.description).slice(0, 3);
      if (!kws.length) return;
      const sameName = norm(target.name).trim();
      cats = cats.map((c) =>
        norm(c.name).trim() === sameName && c.type === target.type
          ? { ...c, keywords: [...new Set([...(c.keywords || []), ...kws])] }
          : c
      );
    });
    const txs = valid.map((d) => ({
      id: uid(), type: d.type, amount: d.amount, description: d.description,
      categoryId: d.categoryId || null, accountId: d.accountId, date: d.date,
    }));
    onSave(txs, cats);
  }

  /* ---------- upload ---------- */
  if (phase === "upload") {
    return (
      <div className="cc-overlay" onClick={onClose}>
        <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
          <div className="cc-grip" />
          <div className="cc-sheet-top">
            <h2>Importar desde Excel</h2>
            <button className="cc-sheet-close" onClick={onClose}>×</button>
          </div>
          <p style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 16 }}>
            Sube un archivo de Excel (.xlsx, .xls) o CSV con tus movimientos. La IA detecta las columnas y los importa.
          </p>

          {config.accounts.length > 1 && (
            <div style={{ marginBottom: 14 }}>
              <label className="cc-label">Cuenta destino</label>
              <select className="cc-select" value={defaultAccountId} onChange={(e) => setDefaultAccountId(e.target.value)}
                style={{ borderColor: !defaultAccountId ? "var(--gold)" : "var(--line)" }}>
                <option value="">Elegir cuenta…</option>
                {config.accounts.map((a) => <option key={a.id} value={a.id}>{a.name}</option>)}
              </select>
            </div>
          )}

          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }}
            onChange={(e) => handleFile(e.target.files?.[0])} />

          <button className="cc-card" onClick={() => inputRef.current?.click()}
            style={{ width: "100%", padding: 26, textAlign: "center", border: "2px dashed var(--line)",
              background: "var(--surface-2)", cursor: "pointer", marginBottom: 14 }}>
            <div style={{ fontSize: 30, marginBottom: 6 }}>📄</div>
            <div style={{ fontWeight: 600, fontSize: 14 }}>
              {file ? file.name : "Toca para elegir tu archivo"}
            </div>
            <div style={{ fontSize: 12, color: "var(--ink-soft)", marginTop: 4 }}>
              {file ? `${rows.length} filas detectadas` : "Acepta .xlsx, .xls y .csv"}
            </div>
          </button>

          {error && (
            <div style={{ padding: 11, borderRadius: 10, background: "var(--coral-soft)", color: "var(--coral)", fontSize: 13, marginBottom: 14, fontWeight: 600 }}>
              {error}
            </div>
          )}

          <button className="cc-btn cc-btn-green" style={{ width: "100%", padding: 14, fontSize: 15 }}
            disabled={!file || !rows.length || !defaultAccountId} onClick={process}>
            ✨ Detectar e importar
          </button>
        </div>
      </div>
    );
  }

  /* ---------- processing ---------- */
  if (phase === "processing") {
    return (
      <div className="cc-overlay">
        <div className="cc-sheet" style={{ padding: "40px 18px" }}>
          <div className="cc-grip" />
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <div style={{ fontSize: 36, marginBottom: 14 }}>🔍</div>
            <div className="cc-serif" style={{ fontSize: 20, fontWeight: 600, marginBottom: 6 }}>
              Analizando tu archivo…
            </div>
            <div style={{ fontSize: 14, color: "var(--ink-soft)", marginBottom: 18 }}>
              La IA está detectando las columnas
            </div>
            <div className="cc-dots"><span /><span /><span /></div>
          </div>
        </div>
      </div>
    );
  }

  /* ---------- review ---------- */
  const accCats = (type) => config.categories.filter((c) => c.accountId === defaultAccountId && c.type === type);

  return (
    <ReviewScreen
      drafts={drafts}
      updateDraft={updateDraft}
      accCats={accCats}
      onBack={() => { setDrafts([]); setPhase("upload"); }}
      onSave={saveAll}
      onClose={onClose}
      sourceLabel="Detecté"
    />
  );
}

/* ----- helpers para Excel ----- */
function parseAmount(v) {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return v;
  const s = String(v).replace(/[$\s]/g, "").replace(/,/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const s = String(v).trim();
  // ISO directo
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  // DD/MM/YYYY o DD-MM-YYYY o DD.MM.YYYY
  const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    let [, dd, mm, yy] = m;
    if (yy.length === 2) yy = "20" + yy;
    return `${yy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`;
  }
  // intentar Date.parse como último recurso
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return null;
}

/* ============= PANTALLA COMPARTIDA: REVISIÓN DE IMPORTACIÓN ============== */
function ReviewScreen({ drafts, updateDraft, accCats, onBack, onSave, onClose, sourceLabel }) {
  const okCount = drafts.filter((d) => d.selected && d.amount > 0).length;
  const dupCount = drafts.filter((d) => d.duplicate).length;
  const dupSelected = drafts.filter((d) => d.duplicate && d.selected).length;

  const markAllDupsOff = () => {
    drafts.forEach((d) => { if (d.duplicate && d.selected) updateDraft(d.tempId, { selected: false }); });
  };
  const markAllDupsOn = () => {
    drafts.forEach((d) => { if (d.duplicate && !d.selected) updateDraft(d.tempId, { selected: true }); });
  };

  return (
    <div className="cc-overlay" onClick={onClose}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>Revisa los movimientos</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 12 }}>
          {sourceLabel} <b>{drafts.length}</b>. Edita lo que esté mal o desmarca los que no quieras importar.
        </p>

        {dupCount > 0 && (
          <div style={{ display: "flex", alignItems: "center", gap: 9, padding: "11px 13px",
            background: "#FFF6E0", border: "1px solid #E8C97A", borderRadius: 12, marginBottom: 14, fontSize: 13 }}>
            <span style={{ fontSize: 17 }}>⚠️</span>
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 700, marginBottom: 2 }}>
                Detecté {dupCount} posible{dupCount === 1 ? "" : "s"} duplicado{dupCount === 1 ? "" : "s"}
              </div>
              <div style={{ color: "var(--ink-soft)", fontSize: 12 }}>
                Movimientos que ya tienes registrados. Están desmarcados por defecto.
              </div>
            </div>
            <button className="cc-btn" style={{ padding: "5px 10px", fontSize: 11.5, whiteSpace: "nowrap" }}
              onClick={dupSelected === dupCount ? markAllDupsOff : markAllDupsOn}>
              {dupSelected === dupCount ? "Desmarcar todos" : "Marcar todos"}
            </button>
          </div>
        )}

        <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 18 }}>
          {drafts.map((d) => {
            const isDup = !!d.duplicate;
            return (
              <div key={d.tempId} className="cc-card" style={{
                padding: 12, opacity: d.selected ? 1 : 0.5,
                borderColor: isDup ? "#E8C97A" : (!d.categoryId ? "var(--gold)" : "var(--line)"),
                background: isDup ? "#FFFCF3" : "var(--surface)",
              }}>
                {isDup && (
                  <div style={{ fontSize: 11.5, fontWeight: 700, color: "#9A6B16",
                    background: "#FFF0CC", padding: "3px 8px", borderRadius: 6,
                    display: "inline-block", marginBottom: 8, letterSpacing: ".02em" }}>
                    🔁 YA EXISTE · {d.duplicate.date}{d.duplicate.description ? ` · "${d.duplicate.description}"` : ""}
                  </div>
                )}
                <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 9 }}>
                  <input type="checkbox" checked={d.selected}
                    onChange={(e) => updateDraft(d.tempId, { selected: e.target.checked })}
                    style={{ width: 18, height: 18, accentColor: "var(--green)" }} />
                  <input className="cc-input" value={d.description} placeholder="Concepto"
                    onChange={(e) => updateDraft(d.tempId, { description: e.target.value })}
                    style={{ flex: 1, fontSize: 14, padding: "7px 10px", fontWeight: 600 }} />
                  <input className="cc-input cc-num" type="number" value={d.amount}
                    onChange={(e) => updateDraft(d.tempId, { amount: parseFloat(e.target.value) || 0 })}
                    style={{ width: 95, fontSize: 14, padding: "7px 10px", fontWeight: 700,
                      color: d.type === "income" ? "var(--green)" : "var(--coral)" }} />
                </div>
                <div style={{ display: "flex", gap: 7, flexWrap: "wrap" }}>
                  <select className="cc-select" value={d.type}
                    onChange={(e) => updateDraft(d.tempId, { type: e.target.value, categoryId: "" })}
                    style={{ flex: "0 0 100px", fontSize: 13, padding: "6px 9px" }}>
                    <option value="expense">Gasto</option>
                    <option value="income">Ingreso</option>
                  </select>
                  <input className="cc-input" type="date" value={d.date}
                    onChange={(e) => updateDraft(d.tempId, { date: e.target.value })}
                    style={{ flex: "0 0 130px", fontSize: 13, padding: "6px 9px" }} />
                  <select className="cc-select" value={d.categoryId}
                    onChange={(e) => updateDraft(d.tempId, { categoryId: e.target.value })}
                    style={{ flex: 1, fontSize: 13, padding: "6px 9px",
                      borderColor: !d.categoryId ? "var(--gold)" : "var(--line)" }}>
                    <option value="">⚠️ Elige categoría</option>
                    {accCats(d.type).map((c) => (
                      <option key={c.id} value={c.id}>{c.emoji} {c.name}</option>
                    ))}
                  </select>
                </div>
              </div>
            );
          })}
        </div>

        <div style={{ display: "flex", gap: 8 }}>
          <button className="cc-btn" style={{ flex: "0 0 auto", padding: 12 }} onClick={onBack}>← Volver</button>
          <button className="cc-btn cc-btn-green" style={{ flex: 1, padding: 13, fontSize: 14 }}
            disabled={okCount === 0} onClick={onSave}>
            Guardar {okCount > 0 ? `(${okCount})` : ""}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============= MODAL: ELEGIR MOVIMIENTO A VINCULAR ======================= */
function LinkPickerModal({ config, txs, currentType, currentAccountId, excludeIds, initialSelectedIds, onClose, onPick }) {
  const [query, setQuery] = useState("");
  const [onlyOpposite, setOnlyOpposite] = useState(false); // por defecto OFF: permitir ambos tipos en el grupo
  const [onlySameAccount, setOnlySameAccount] = useState(true);
  const [selected, setSelected] = useState(new Set(initialSelectedIds || []));

  const exclude = new Set(excludeIds || []);
  const opposite = currentType === "income" ? "expense" : "income";

  // filtrar
  let list = txs.filter((t) => !exclude.has(t.id));
  if (onlyOpposite) list = list.filter((t) => t.type === opposite);
  if (onlySameAccount && currentAccountId) list = list.filter((t) => t.accountId === currentAccountId);
  if (query.trim()) {
    const q = norm(query.trim());
    list = list.filter((t) => {
      const desc = norm(t.description || "");
      const amt = String(t.amount);
      const cat = config.categories.find((c) => c.id === t.categoryId);
      const catName = cat ? norm(cat.name) : "";
      return desc.includes(q) || amt.includes(q) || catName.includes(q) || t.date.includes(query.trim());
    });
  }
  list = [...list].sort((a, b) => b.date.localeCompare(a.date) || (b.id || "").localeCompare(a.id || ""));

  const accMulti = config.accounts.length > 1;
  const toggle = (id) => {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  // resumen en vivo del grupo (incluyendo solo el actual)
  const selectedTxs = Array.from(selected).map((id) => txs.find((t) => t.id === id)).filter(Boolean);
  const totalIn = selectedTxs.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const totalOut = selectedTxs.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);

  return (
    <div className="cc-overlay" onClick={onClose} style={{ zIndex: 70 }}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div className="cc-sheet-top">
          <h2>Elegir movimientos</h2>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>
        <p style={{ fontSize: 13, color: "var(--ink-soft)", marginBottom: 12 }}>
          Marca todos los movimientos que forman parte del mismo reembolso o encargo. Puedes elegir varios ingresos y/o varios gastos.
        </p>

        <div style={{ marginBottom: 10 }}>
          <input className="cc-input"
            placeholder="🔍 Buscar por concepto, monto, fecha…"
            value={query} onChange={(e) => setQuery(e.target.value)} autoFocus />
        </div>

        <div style={{ display: "flex", gap: 7, marginBottom: 12, flexWrap: "wrap" }}>
          <button onClick={() => setOnlyOpposite(!onlyOpposite)}
            className="cc-chip"
            style={{
              background: onlyOpposite ? "var(--ink)" : "var(--surface)",
              color: onlyOpposite ? "var(--surface)" : "var(--ink)",
              borderColor: onlyOpposite ? "var(--ink)" : "var(--line)",
            }}>
            {onlyOpposite ? "✓ " : ""}Solo {opposite === "income" ? "ingresos" : "gastos"}
          </button>
          {accMulti && (
            <button onClick={() => setOnlySameAccount(!onlySameAccount)}
              className="cc-chip"
              style={{
                background: onlySameAccount ? "var(--ink)" : "var(--surface)",
                color: onlySameAccount ? "var(--surface)" : "var(--ink)",
                borderColor: onlySameAccount ? "var(--ink)" : "var(--line)",
              }}>
              {onlySameAccount ? "✓ " : ""}Misma cuenta
            </button>
          )}
          <span style={{ marginLeft: "auto", alignSelf: "center", fontSize: 11.5, color: "var(--ink-soft)", whiteSpace: "nowrap" }}>
            {list.length} resultado{list.length === 1 ? "" : "s"}
          </span>
        </div>

        {/* lista con checkboxes */}
        <div style={{ display: "flex", flexDirection: "column", gap: 0,
          maxHeight: selected.size > 0 ? "40vh" : "55vh", overflowY: "auto",
          border: "1px solid var(--line)", borderRadius: 12, background: "var(--surface)", padding: "4px 14px" }}>
          {list.length === 0 ? (
            <div style={{ padding: "26px 4px", textAlign: "center", color: "var(--ink-soft)", fontSize: 14 }}>
              <div style={{ fontSize: 28, marginBottom: 6 }}>🔍</div>
              No hay movimientos que coincidan.
              {query && <div style={{ fontSize: 12, marginTop: 4 }}>Prueba con otros términos o quita los filtros.</div>}
            </div>
          ) : (
            renderGroupedByDay(list).map((entry) =>
              entry.type === "header" ? (
                <div key={`h-${entry.date}`} className="cc-day-sep">
                  {(() => { const p = dayParts(entry.date); return (
                    <>
                      <span className="cc-day-num">{p.num}</span>
                      <span className="cc-day-name">{p.desc}</span>
                    </>
                  ); })()}
                </div>
              ) : (
                <LinkPickRow key={entry.t.id} t={entry.t} config={config}
                  selected={selected.has(entry.t.id)}
                  onPick={() => toggle(entry.t.id)} />
              )
            )
          )}
        </div>

        {/* resumen en vivo + botón confirmar */}
        {selected.size > 0 && (
          <div style={{ marginTop: 12, padding: "10px 12px", background: "#FFF6E0", border: "1px solid #E8C97A",
            borderRadius: 11, fontSize: 12.5 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
              <span>Ingresos seleccionados:</span>
              <span className="cc-num" style={{ fontWeight: 700, color: "var(--green)" }}>{fmt(totalIn)}</span>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span>Gastos seleccionados:</span>
              <span className="cc-num" style={{ fontWeight: 700, color: "var(--coral)" }}>{fmt(totalOut)}</span>
            </div>
          </div>
        )}

        <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
          {selected.size > 0 && (
            <button className="cc-btn" style={{ padding: "12px 14px", fontSize: 13 }}
              onClick={() => setSelected(new Set())}>
              Limpiar
            </button>
          )}
          <button className="cc-btn cc-btn-green" style={{ flex: 1, padding: 13, fontSize: 14 }}
            disabled={selected.size === 0}
            onClick={() => onPick(Array.from(selected))}>
            {selected.size === 0 ? "Selecciona al menos uno" : `Vincular ${selected.size} ${selected.size === 1 ? "movimiento" : "movimientos"}`}
          </button>
        </div>
      </div>
    </div>
  );
}

function LinkPickRow({ t, config, selected, onPick }) {
  const c = config.categories.find((x) => x.id === t.categoryId);
  const acc = config.accounts.find((a) => a.id === t.accountId);
  const multi = config.accounts.length > 1;
  return (
    <div
      onClick={onPick}
      style={{ display: "flex", alignItems: "center", gap: 11, padding: "11px 0",
        borderBottom: "1px solid var(--line)", cursor: "pointer",
        background: selected ? "var(--green-soft)" : "transparent",
        margin: "0 -10px", paddingLeft: 10, paddingRight: 10,
        borderRadius: selected ? 8 : 0, transition: ".12s" }}
      onMouseEnter={(e) => { if (!selected) e.currentTarget.style.background = "var(--surface-2)"; }}
      onMouseLeave={(e) => { if (!selected) e.currentTarget.style.background = "transparent"; }}>
      <div style={{ fontSize: 22, width: 30, textAlign: "center" }}>{c ? c.emoji : "❔"}</div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ fontWeight: 600, fontSize: 14.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
          {t.description || (c ? c.name : "Movimiento")}
        </div>
        <div style={{ fontSize: 12, color: "var(--ink-soft)" }}>
          {c ? c.name : "Sin categoría"}{multi && acc ? ` · ${acc.name}` : ""}
        </div>
      </div>
      <div className="cc-num" style={{ fontWeight: 700, fontSize: 15, whiteSpace: "nowrap",
        color: t.type === "income" ? "var(--green)" : "var(--coral)" }}>
        {t.type === "income" ? "+" : "−"}{fmt(t.amount).replace("-", "")}
      </div>
      {selected && <span style={{ color: "var(--green)", fontSize: 18, fontWeight: 700 }}>✓</span>}
    </div>
  );
}

/* ============= TARJETA RESUMEN del periodo (pestaña Movimientos) ========= */
function SummaryCard({ filter, totalIn, totalOut, topCatRows, topTotal, config }) {
  const showIn = filter === "all" || filter === "income";
  const showOut = filter === "all" || filter === "expense";
  const net = totalIn - totalOut;
  return (
    <div className="cc-card" style={{ padding: "14px 16px", marginBottom: 12 }}>
      <div style={{ display: "flex", alignItems: "stretch", gap: 12 }}>
        {showIn && (
          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 2 }}>
            <div style={{ fontSize: 10.5, fontWeight: 600, color: "var(--ink-faint)",
              textTransform: "uppercase", letterSpacing: ".06em" }}>Ingresos</div>
            <div className="cc-serif cc-num" style={{ fontSize: 20, fontWeight: 500, color: "var(--ink)" }}>{fmt(totalIn)}</div>
          </div>
        )}
        {showOut && (
          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 2,
            paddingLeft: showIn ? 12 : 0, borderLeft: showIn ? "1px solid var(--line-soft)" : "none" }}>
            <div style={{ fontSize: 10.5, fontWeight: 600, color: "var(--ink-faint)",
              textTransform: "uppercase", letterSpacing: ".06em" }}>Gastos</div>
            <div className="cc-serif cc-num" style={{ fontSize: 20, fontWeight: 500, color: "var(--coral)" }}>{fmt(totalOut)}</div>
          </div>
        )}
        {filter === "all" && (
          <div style={{ flex: 1, display: "flex", flexDirection: "column", gap: 2,
            paddingLeft: 12, borderLeft: "1px solid var(--line-soft)" }}>
            <div style={{ fontSize: 10.5, fontWeight: 600, color: "var(--ink-faint)",
              textTransform: "uppercase", letterSpacing: ".06em" }}>Flujo neto</div>
            <div className="cc-serif cc-num" style={{ fontSize: 20, fontWeight: 500,
              color: net >= 0 ? "var(--ink)" : "var(--coral)" }}>
              {net >= 0 ? "+" : "−"}{fmt(Math.abs(net))}
            </div>
          </div>
        )}
      </div>

      {topCatRows.length > 0 && (
        <div style={{ marginTop: 10, paddingTop: 10, borderTop: "1px solid var(--line)",
          display: "flex", flexWrap: "wrap", gap: 8, fontSize: 11.5 }}>
          {topCatRows.map(([id, amt]) => {
            const c = config.categories.find((x) => x.id === id);
            const pct = topTotal ? Math.round((amt / topTotal) * 100) : 0;
            return (
              <div key={id} style={{ display: "inline-flex", alignItems: "center", gap: 5,
                background: "var(--surface-2)", padding: "4px 9px", borderRadius: 99 }}>
                <span style={{ fontSize: 13 }}>{c ? c.emoji : "❔"}</span>
                <span style={{ fontWeight: 600 }}>{c ? c.name : "Sin cat"}</span>
                <span style={{ color: "var(--ink-soft)" }}>{pct}%</span>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* ============= MODAL DETALLE: drill-down de ingresos/gastos/categoría ==== */
function DetailModal({ config, detail, dateRange, onClose, onEditTx }) {
  // ordenar más recientes primero, agrupar por día
  const list = [...(detail.txs || [])].sort((a, b) => b.date.localeCompare(a.date) || (b.id || "").localeCompare(a.id || ""));
  const count = list.length;

  return (
    <div className="cc-overlay" onClick={onClose} style={{ zIndex: 60 }}>
      <div className="cc-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="cc-grip" />
        <div style={{ display: "flex", alignItems: "flex-start", gap: 12, marginBottom: 14 }}>
          <div className="cc-emoji" style={{ fontSize: 32, lineHeight: 1, marginTop: 2 }}>{detail.emoji}</div>
          <div style={{ flex: 1, minWidth: 0 }}>
            <h2 className="cc-serif" style={{ fontSize: 21, fontWeight: 600, marginBottom: 2 }}>{detail.label}</h2>
            <div style={{ fontSize: 12.5, color: "var(--ink-soft)" }}>
              {rangeLabel(dateRange)} · {count} movimiento{count === 1 ? "" : "s"}
            </div>
          </div>
          <button className="cc-sheet-close" onClick={onClose}>×</button>
        </div>

        {/* total */}
        <div className="cc-card" style={{ padding: "14px 16px", marginBottom: 14,
          background: detail.kind === "income" ? "var(--green-soft)" : detail.kind === "expense" ? "var(--coral-soft)" : "var(--surface)" }}>
          <div className="cc-label" style={{ marginBottom: 3 }}>Total</div>
          <div className="cc-num" style={{ fontSize: 28, fontWeight: 700, color: detail.color, letterSpacing: "-.01em" }}>
            {fmt(detail.total)}
          </div>
        </div>

        {/* lista de movimientos */}
        <div style={{ maxHeight: "55vh", overflowY: "auto", border: "1px solid var(--line)", borderRadius: 12,
          background: "var(--surface)", padding: "4px 14px" }}>
          {list.length === 0 ? (
            <div style={{ padding: "26px 4px", textAlign: "center", color: "var(--ink-soft)", fontSize: 14 }}>
              Sin movimientos en este filtro.
            </div>
          ) : (
            renderGroupedByDay(list).map((entry) =>
              entry.type === "header" ? (
                <div key={`h-${entry.date}`} className="cc-day-sep">
                  {(() => { const p = dayParts(entry.date); return (
                    <>
                      <span className="cc-day-num">{p.num}</span>
                      <span className="cc-day-name">{p.desc}</span>
                    </>
                  ); })()}
                  <div className="cc-day-totals">
                    {entry.income > 0 && <span className="pos">+{fmtBare(entry.income)}</span>}
                    {entry.expense > 0 && <span className="neg">−{fmtBare(entry.expense)}</span>}
                  </div>
                </div>
              ) : (
                <TxRow key={entry.t.id} t={entry.t} config={config}
                  onEdit={entry.t.id && !entry.t.id.startsWith("synth-") ? onEditTx : null} />
              )
            )
          )}
        </div>
      </div>
    </div>
  );
}
