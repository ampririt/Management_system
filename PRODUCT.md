# Product

## Register

product

## Users

A single owner-operator running a Japan→Vietnam resale/arbitrage side-business alongside their personal finances. They use this at a desk in focused sessions (sometimes on a phone), entering income, expenses, purchase orders, credit-card activity, and currency conversions, then watching for the right moment to exchange their VND wallet back to yen. They are financially literate but not an accountant; they want to trust the numbers at a glance, not audit them.

## Product Purpose

ProfitIntelligence is a single-owner financial command center backed by Google Sheets (Apps Script web app). Goods are bought in ¥ and sold in ₫ at a fixed rate, so the trade is break-even by design; **all profit comes from exchanging the VND wallet back to yen at a better-than-fixed rate.** The app surfaces two wallets (¥ cash, ₫ held), credit-card exposure, a purchase-order pipeline, daily cash flow, and an FX trading view that shows profit per exchange. Success is the owner knowing, in seconds, how much cash they hold, how much VND is waiting to be converted, and whether now is a good time to convert.

## Brand Personality

Calm, trustworthy, quietly confident. Bank-like rather than start-up: muted surfaces, careful typography, money rendered in tabular figures you can scan and trust. Nothing shouts; the numbers are the hero. Voice is plain and direct, never gamified or jargon-heavy.

## Anti-references

- **Generic SaaS-indigo dashboard** — Inter + #4f46e5 + cream `#fafaf7`. This is the current skin and the thing to move away from; it reads "AI made this."
- **Cluttered enterprise ERP** — SAP/Oracle gray-on-gray density, competing toolbars, everything the same weight.
- **Playful consumer fintech** — Revolut/Cash-App gradients, big rounded bubbles, emoji-forward, gamified streaks.
- **Bare spreadsheet** — looks like the Google Sheet underneath; grids of numbers with no hierarchy.

## Design Principles

1. **The number is the hero.** Money is rendered in tabular figures with clear hierarchy; everything else (chrome, labels, icons) recedes so values are scannable at a glance.
2. **Quiet confidence over decoration.** Restraint signals trust for a finance tool. One considered accent, muted surfaces, generous breathing room. No gradients-as-decoration, no glass, no side-stripes.
3. **Answer the owner's question fast.** Every screen leads with the one thing they opened it to learn (cash held, VND waiting, is it a good time to convert), then supporting detail.
4. **Fast feels trustworthy.** Data entry must feel instant; a tool that freezes on every save feels fragile with money. Optimistic updates and perceived-speed polish over full reloads.
5. **Honest, plain language.** Labels say what a value is; buttons say what will happen. No marketing voice, no false precision.

## Accessibility & Inclusion

Target WCAG AA: body text ≥4.5:1, large/secondary text ≥3:1, visible keyboard focus, `prefers-reduced-motion` respected, and state never conveyed by color alone (pair an icon or label). Comfortable for long focused sessions.
