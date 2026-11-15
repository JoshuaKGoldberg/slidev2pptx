<h1 align="center">Slidev2pptx</h1>

<p align="center">Builds native, editable PowerPoint files from Slidev decks.</p>

<p align="center">
	<!-- prettier-ignore-start -->
	<!-- ALL-CONTRIBUTORS-BADGE:START - Do not remove or modify this section -->
	<a href="#contributors" target="_blank"><img alt="👪 All Contributors: 1" src="https://img.shields.io/badge/%F0%9F%91%AA_all_contributors-1-21bb42.svg" /></a>
<!-- ALL-CONTRIBUTORS-BADGE:END -->
	<!-- prettier-ignore-end -->
	<a href="https://github.com/JoshuaKGoldberg/slidev2pptx/blob/main/.github/CODE_OF_CONDUCT.md" target="_blank"><img alt="🤝 Code of Conduct: Kept" src="https://img.shields.io/badge/%F0%9F%A4%9D_code_of_conduct-kept-21bb42" /></a>
	<a href="https://codecov.io/gh/JoshuaKGoldberg/slidev2pptx" target="_blank"><img alt="🧪 Coverage" src="https://img.shields.io/codecov/c/github/JoshuaKGoldberg/slidev2pptx?label=%F0%9F%A7%AA%20coverage" /></a>
	<a href="https://github.com/JoshuaKGoldberg/slidev2pptx/blob/main/LICENSE.md" target="_blank"><img alt="📝 License: MIT" src="https://img.shields.io/badge/%F0%9F%93%9D_license-MIT-21bb42.svg" /></a>
	<a href="http://npmjs.com/package/slidev2pptx" target="_blank"><img alt="📦 npm version" src="https://img.shields.io/npm/v/slidev2pptx?color=21bb42&label=%F0%9F%93%A6%20npm" /></a>
	<img alt="💪 TypeScript: Strict" src="https://img.shields.io/badge/%F0%9F%92%AA_typescript-strict-21bb42.svg" />
</p>

## Usage

```shell
npm i -D slidev2pptx
```

Run it next to a Slidev deck to write a native, editable PowerPoint file:

```shell
npx slidev2pptx slides.md --output talk.pptx
```

Text, lists, shapes, and pictures become native PowerPoint objects, and `v-click`s become click animations.
Anything that can't be exported as it looks in Slidev is reported as a warning or error.
Run `npx slidev2pptx --help` for more options, such as building a range of slides or one slide per click.

It also exports a `buildPptx` function:

```ts
import { buildPptx } from "slidev2pptx";

const { output, problems } = await buildPptx({
	entry: "slides.md",
	output: "talk.pptx",
});
```

## Requirements

- Node.js `^22.22.0 || >=24.0.0`
- Slidev `^53.1.0`: slidev2pptx renders slides with Slidev's own client internals, so it supports one Slidev major version at a time
- The deck's fonts installed on the machine that opens the PowerPoint file, as fonts aren't embedded in it.
  slidev2pptx lays text out with the deck's `@font-face` fonts, its Slidev web fonts, and installed fonts, and warns about any it can't find.

## Supported Features

- Text, lists, code blocks, boxes, borders, and SVG shapes become native, editable PowerPoint shapes, with links kept clickable
- Pictures are kept as pictures, and math is drawn as a picture whose alt text is its TeX source
- `v-click`s and other click states become entrance and exit animations, or one slide per click state with `--with-clicks`
- Slide transitions become PowerPoint's closest transition: `slide-*` as a push, anything else as a fade
- Speaker notes, dark mode with `--dark`, and slide ranges with `--range`

## Known Limitations

- Mermaid diagrams, embedded posts, `<iframe>`s, and `<video>`s aren't exported, and are reported as errors
- The output of `monaco-run` blocks and the modified side of `monaco-diff` blocks aren't exported, and are reported as errors
- `v-mark` annotations aren't drawn, and are reported as warnings
- An element whose contents would draw more than 400 shapes, such as a very dense slide, is drawn as one picture instead of editable shapes
- PowerPoint can't number list items below 1, so such lists are numbered from 1, with a warning
- PowerPoint lays text out itself, so lines can wrap or sit slightly differently than in Slidev.
  For example, PowerPoint for Mac sometimes leaves out a font's pair kerning, making lines a few percent wider, and paragraphs that mix font sizes can drift by a few pixels.
  [`qa/reports/`](./qa/reports/) lists the known differences.

## Development

See [`.github/CONTRIBUTING.md`](./.github/CONTRIBUTING.md), then [`.github/DEVELOPMENT.md`](./.github/DEVELOPMENT.md).
Thanks! 💖

## Contributors

<!-- spellchecker: disable -->
<!-- ALL-CONTRIBUTORS-LIST:START - Do not remove or modify this section -->
<!-- prettier-ignore-start -->
<!-- markdownlint-disable -->
<table>
  <tbody>
    <tr>
      <td align="center"><a href="http://www.joshuakgoldberg.com"><img src="https://avatars.githubusercontent.com/u/3335181?v=4?s=100" width="100px;" alt="Josh Ghoulberg 👻"/><br /><sub><b>Josh Ghoulberg 👻</b></sub></a><br /><a href="https://github.com/JoshuaKGoldberg/slidev2pptx/commits?author=JoshuaKGoldberg" title="Code">💻</a> <a href="#content-JoshuaKGoldberg" title="Content">🖋</a> <a href="https://github.com/JoshuaKGoldberg/slidev2pptx/commits?author=JoshuaKGoldberg" title="Documentation">📖</a> <a href="#ideas-JoshuaKGoldberg" title="Ideas, Planning, & Feedback">🤔</a> <a href="#infra-JoshuaKGoldberg" title="Infrastructure (Hosting, Build-Tools, etc)">🚇</a> <a href="#maintenance-JoshuaKGoldberg" title="Maintenance">🚧</a> <a href="#projectManagement-JoshuaKGoldberg" title="Project Management">📆</a> <a href="#tool-JoshuaKGoldberg" title="Tools">🔧</a></td>
    </tr>
  </tbody>
</table>

<!-- markdownlint-restore -->
<!-- prettier-ignore-end -->

<!-- ALL-CONTRIBUTORS-LIST:END -->
<!-- spellchecker: enable -->

> 💝 This package was templated with [`create-typescript-app`](https://github.com/JoshuaKGoldberg/create-typescript-app) using the [Bingo framework](https://create.bingo).
