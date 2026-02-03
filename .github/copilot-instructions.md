# Copilot Instructions

## Project Overview

This workspace contains two interactive HTML/CSS/JavaScript applications:

1. **index.html** – Digital Strategy presentation (5-slide carousel with transitions, animations, and keyboard navigation)
2. **TFL** – Tigres Asiáticos Travel website (responsive landing page for Southeast Asia tour packages)

Both are single-file, self-contained applications with embedded styling and scripting. No build tools, frameworks, or external dependencies—pure vanilla HTML/CSS/JS.

## Architecture & Patterns

### Single-File Philosophy
- **All-in-one structure**: CSS in `<style>`, JS in `<script>` tags, no external files
- **CSS Variables for theming**: Root-level color system (`--bg0`, `--text`, `--accent`, etc.)
- **Semantic HTML**: Proper ARIA roles, labels, and landmarks for accessibility
- **No frameworks**: Pure DOM manipulation and vanilla JavaScript

### Key Patterns

**CSS Architecture (Both Projects)**
- **Variable-driven design**: All colors, spacing, shadows defined in `:root`
- **Responsive breakpoints**: `max-width: 980px` (tablet), `max-width: 520px` (mobile)
- **Glassmorphism/gradient design**: Backdrop filters, radial gradients, conic gradients
- **Utility-like approach**: Reusable classes (`.card`, `.btn`, `.tag`, `.grid`, etc.)

**JavaScript Conventions**
- **Functional over OOP**: Event listeners, selectors via `$()` and `$$()` shortcuts
- **State management**: Minimal—track current slide index or filter state directly
- **Smooth interactions**: CSS transitions + JS for state updates (no animations libraries)
- **Progressive enhancement**: Form fallbacks, keyboard/touch/mouse support

## Development Workflow

### Testing
- No test framework—visually inspect in browser
- Responsiveness: Check breakpoints (980px, 520px)
- Accessibility: Verify ARIA attributes, keyboard navigation (arrows, space, tab)

### Common Edits
- **Colors**: Modify `:root` variables
- **Typography**: Adjust `font-size` clamps and `line-height` 
- **Spacing**: Update `gap`, `padding`, `margin` in layout classes
- **Content**: Edit HTML text, form fields, or card data
- **Interactions**: Add event listeners or modify smooth scroll behavior

### Performance Notes
- **Animations**: Use CSS `transition` + `transform` (GPU-accelerated)
- **Avoid**: `left`/`top` for movement; use `translateX`/`translateY`
- **Lazy loading**: Not needed for single-page sites; use `defer` on scripts if any

## File-Specific Details

### index.html (Presentation)
- **Slide engine**: Array-based carousel with transform-X navigation
- **Progress bar**: `scaleX()` based on current slide index
- **Animations**: Per-slide intro with staggered delays (`.anim.d1`, `.d2`, etc.)
- **Controls**: Previous/Next buttons, dot indicators, keyboard shortcuts (arrows, space, home/end)
- **SVG diagrams**: Inline `<svg>` for journey funnel, data pipeline, operating model

### TFL (Travel Website)
- **Hero + sections**: Navigation, hero form, destinations grid, itineraries, FAQ, footer
- **Form prefill**: Click itinerary button → prepopulates contact form with selected trip
- **Filter system**: Chip buttons filter itinerary cards by data-tags
- **Mobile drawer**: Side menu triggered by hamburger button
- **Reveal animation**: Intersection observer triggers staggered `.in` class on scroll

## Conventions to Follow

1. **Color usage**: Always reference `:root` variables, never hardcode hex/rgba
2. **Spacing scale**: Use `gap`, `padding`, `margin` consistently (8px, 12px, 14px, 16px, 18px multiples)
3. **Border radius**: Prefer `--r-xl`, `--r-lg`, `--r-md` or `999px` for pills
4. **Typography sizing**: Use `clamp()` for responsive sizes (e.g., `clamp(24px, 2.6vw, 36px)`)
5. **Form fields**: All inputs use `.field` class with consistent border, background, focus states
6. **Buttons**: `.btn`, `.btn.primary`, `.btn.ghost` classes—avoid inline button styling
7. **Grid layouts**: Use CSS Grid with `.grid.two`, `.grid.three`, `.grid.six` for responsive columns

## Quick Commands

- **Open in browser**: Simply open the `.html` file in any modern browser (Chrome, Firefox, Safari, Edge)
- **Live preview**: Use VS Code's Live Server extension or any local HTTP server
- **Accessibility check**: Use browser DevTools accessibility tree or axe DevTools extension

## Integration Points

- **Contact forms**: Currently demo (console.log via `alert()`); integrate with backend API or CRM via `fetch()` in form submit handler
- **Analytics**: No GA/tracking code present—add `<script>` tag for analytics library if needed
- **CMS integration**: Content is hardcoded; for dynamic content, add API calls in JS or use server-side templating

---

**Generated for AI agents working in this codebase. Focus on vanilla HTML/CSS/JS patterns, semantic markup, and responsive design.**
