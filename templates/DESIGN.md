---
name: Gongju NGPP Industrial Control System
colors:
  surface: '#f8f9ff'
  surface-dim: '#cbdbf5'
  surface-bright: '#f8f9ff'
  surface-container-lowest: '#ffffff'
  surface-container-low: '#eff4ff'
  surface-container: '#e5eeff'
  surface-container-high: '#dce9ff'
  surface-container-highest: '#d3e4fe'
  on-surface: '#0b1c30'
  on-surface-variant: '#44474c'
  inverse-surface: '#213145'
  inverse-on-surface: '#eaf1ff'
  outline: '#74777d'
  outline-variant: '#c4c6cd'
  surface-tint: '#4f6073'
  primary: '#041627'
  on-primary: '#ffffff'
  primary-container: '#1a2b3c'
  on-primary-container: '#8192a7'
  inverse-primary: '#b7c8de'
  secondary: '#436180'
  on-secondary: '#ffffff'
  secondary-container: '#bcdafe'
  on-secondary-container: '#42607f'
  tertiary: '#001818'
  on-tertiary: '#ffffff'
  tertiary-container: '#002f2f'
  on-tertiary-container: '#3a9f9e'
  error: '#ba1a1a'
  on-error: '#ffffff'
  error-container: '#ffdad6'
  on-error-container: '#93000a'
  primary-fixed: '#d2e4fb'
  primary-fixed-dim: '#b7c8de'
  on-primary-fixed: '#0b1d2d'
  on-primary-fixed-variant: '#38485a'
  secondary-fixed: '#d0e4ff'
  secondary-fixed-dim: '#abc9ed'
  on-secondary-fixed: '#001d34'
  on-secondary-fixed-variant: '#2b4967'
  tertiary-fixed: '#93f2f2'
  tertiary-fixed-dim: '#76d6d5'
  on-tertiary-fixed: '#002020'
  on-tertiary-fixed-variant: '#004f4f'
  background: '#f8f9ff'
  on-background: '#0b1c30'
  surface-variant: '#d3e4fe'
typography:
  display-lg:
    fontFamily: Hanken Grotesk
    fontSize: 32px
    fontWeight: '700'
    lineHeight: '1.2'
  headline-md:
    fontFamily: Hanken Grotesk
    fontSize: 24px
    fontWeight: '600'
    lineHeight: '1.3'
  title-sm:
    fontFamily: Hanken Grotesk
    fontSize: 18px
    fontWeight: '600'
    lineHeight: '1.4'
  body-base:
    fontFamily: Pretendard
    fontSize: 15px
    fontWeight: '400'
    lineHeight: '1.6'
  body-sm:
    fontFamily: Pretendard
    fontSize: 13px
    fontWeight: '400'
    lineHeight: '1.5'
  data-mono:
    fontFamily: JetBrains Mono
    fontSize: 13px
    fontWeight: '500'
    lineHeight: '1.4'
    letterSpacing: -0.02em
  label-caps:
    fontFamily: Hanken Grotesk
    fontSize: 11px
    fontWeight: '700'
    lineHeight: '1'
    letterSpacing: 0.05em
rounded:
  sm: 0.125rem
  DEFAULT: 0.25rem
  md: 0.375rem
  lg: 0.5rem
  xl: 0.75rem
  full: 9999px
spacing:
  base: 4px
  xs: 4px
  sm: 8px
  md: 16px
  lg: 24px
  xl: 40px
  container-max: 1600px
  sidebar-width: 260px
---

## Brand & Style

This design system is engineered for the high-stakes environment of the Gongju Natural Gas Power Plant. The brand personality is **authoritative, systematic, and resilient**. It prioritizes cognitive clarity over decorative flair, ensuring that plant operators can monitor complex telemetry and manage workflows without visual fatigue.

The design style follows a **Corporate Modern** approach with a **Technical/Utilitarian** edge. It utilizes a structured information architecture characterized by:
- **High Density:** Efficient use of screen real estate to minimize scrolling in data-heavy views.
- **Visual Precision:** Micro-interactions that provide immediate, unambiguous feedback.
- **Reliability:** A stable, predictable layout that evokes a sense of security and institutional permanence.
- **Korean Optimization:** Specialized focus on CJK legibility, ensuring high-stroke-density characters remain clear in dense tables.

## Colors

The palette is anchored in **Deep Navy (#1A2B3C)** to establish a foundation of stability and professional rigor. 

- **Primary & Secondary:** Used for navigation backgrounds, primary actions, and structural headers. The slate blue shifts provide subtle hierarchy within the sidebar and masthead.
- **Accents (Teal & Amber):** Teal (#008080) is used for active states and "Verified" status to provide a high-contrast alternative to standard green. Amber (#D97706) is reserved for "Pending" or "Cautionary" states.
- **Neutrals:** A range of Cool Slates are used for borders, secondary text, and background layering to prevent the UI from feeling "flat" while maintaining a professional gray-scale balance.
- **Surface:** Backgrounds use a very light off-white/gray (#F8FAFC) to reduce glare during long shifts.

## Typography

Typography is optimized for **Multilingual Clarity (Korean/English)** and **Numeric Readability**.

- **Primary Sans:** Hanken Grotesk provides a modern, sharp geometric feel for headings and UI controls.
- **Body Text:** Pretendard (or Inter as a fallback) is specified for its exceptional Korean character balance and legibility in small-scale data tables.
- **Technical Mono:** JetBrains Mono is used exclusively for sensor readings, timestamps, and ID numbers, ensuring that digits are tabular and easily comparable at a glance.
- **Hierarchy:** Use bold weights sparingly to highlight critical status changes. Ensure line heights for Korean text are slightly more generous than standard Latin settings to accommodate complex glyphs.

## Layout & Spacing

The layout utilizes a **Structured Fluid Grid** optimized for 1080p and 1440p industrial monitors.

- **Grid System:** A 12-column grid with 16px (md) gutters. On desktop, the sidebar is fixed, while the content area expands.
- **Density:** This system employs a "Compact" spacing model. Vertical padding in tables is kept to 8px-12px to maximize data visibility.
- **Breakpoints:**
  - **Desktop (1280px+):** Full sidebar and expanded data visualization.
  - **Tablet (768px - 1279px):** Collapsed icon-only sidebar; tables switch to horizontal scroll.
  - **Mobile:** Not recommended for full control; used for status alerts and read-only summaries.

## Elevation & Depth

To maintain a "Professional/Secure" feel, this design system avoids heavy shadows in favor of **Tonal Layering and Low-Contrast Outlines**.

- **Surface Tiers:**
  - **Level 0 (Background):** #F8FAFC (Light Slate)
  - **Level 1 (Cards/Tables):** #FFFFFF (Pure White) with 1px border #E2E8F0.
  - **Level 2 (Modals/Popovers):** #FFFFFF with a crisp 4px blur, 10% opacity black shadow to lift it from the work surface.
- **Interaction Depth:** Instead of neomorphic extrusions, use "inset" states (subtle 1px inner shadows) for active button presses to simulate physical tactile switches.
- **Dividers:** Use #E2E8F0 for standard separation. Use #CBD5E1 for section-level headers.

## Shapes

The shape language is **Soft (0.25rem)**. This subtle rounding removes the aggressive "sharpness" of pure industrial software while maintaining a structured, architectural feel.

- **Buttons/Inputs:** 4px radius.
- **Large Containers/Modals:** 8px (rounded-lg) to provide a clear visual container for complex forms.
- **Status Badges:** 2px or fully square to differentiate them from interactive elements.
- **Iconography:** Use a consistent 2px stroke weight with slightly rounded joins to match the component radius.

## Components

### Data Tables
- **Header:** Sticky headers with a Primary Navy background and White text.
- **Rows:** Zebra-striping every second row (#F1F5F9).
- **Cells:** Numeric data must be right-aligned and monospaced. Status badges must be centered.

### Status Badges
- **Verified:** Teal background (10% opacity) with Teal bold text.
- **Error:** Red background (10% opacity) with Red bold text and a leading "!" icon.
- **Pending:** Amber background (10% opacity) with Amber bold text.

### Buttons
- **Primary:** Solid Deep Navy background.
- **Secondary:** Transparent with a Slate Blue border.
- **Tertiary:** Text-only with an underline on hover.

### File Upload Zones
- Dashed border #CBD5E1. High-contrast drag-and-drop state using a Teal tint. Include a file-type icon (PDF/XLS/DWG) in the preview.

### Input Fields
- Labels are always positioned above the field in `label-caps` style.
- Active state uses a 2px Teal border. Error state uses a 2px Red border with helper text below.

### Modals
- Focused on "Detail Views." Always include a clear "X" close button and a secondary "Cancel" button in the footer to prevent accidental data loss.