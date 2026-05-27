# Frontend Design Skill

## Typography
- Use Inter or fitting Google Font, never system default
- Sizes: 14px body, 16px emphasis, 24-32px headings
- Line height: 1.5 body, 1.2 headings
- font-medium (500) for emphasis

## Color
- NEVER use #0000ff blue or #800080 purple
- 1 primary, 1 secondary, 1 accent color
- Grays: slate-50, slate-100, slate-800, slate-900
- Hover: darken 10%, same hue

## Spacing
- Scale: 4, 8, 12, 16, 24, 32, 48, 64
- Cards: p-6
- Between sections: gap-8 or gap-12
- Related items: gap-4
- No arbitrary values (p-7, gap-5)

## Layout
- Max width: 1280px (max-w-7xl)
- Grid for layouts, Flex for alignment
- Cards: shadow-sm or shadow-md
- Corners: rounded-lg cards, rounded-md buttons

## Don'ts
- No gradients unless requested
- No animations on everything
- No pure black - use slate-900
- No low-contrast text

## Example
```tsx
<div className="bg-white rounded-lg shadow-sm border border-slate-200 p-6">
  <h3 className="text-lg font-medium text-slate-900">Title</h3>
  <p className="mt-2 text-sm text-slate-600">Description</p>
  <button className="mt-4 w-full bg-indigo-600 hover:bg-indigo-700 text-white font-medium py-2.5 px-4 rounded-md transition-colors">
    Action
  </button>
</div>
```
