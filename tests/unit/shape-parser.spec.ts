import { describe, expect, it } from 'vitest';
import { parseShape, parseShapeProperties, parseTransform2D } from '../../src/ooxml/drawingml/parsing/shape-parser';
import type { OpenXmlElement } from '../../src/ooxml/wordprocessingml/model/element';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

const MOCK_OPTIONS = { debug: false } as any;
const MOCK_CALLBACKS = { parseBodyElements: () => [] };

function makeWsp(inner: string): Element {
    const xml = `<wps:wsp xmlns:wps="${WPS_NS}" xmlns:a="${A_NS}">${inner}</wps:wsp>`;
    return new DOMParser().parseFromString(xml, 'text/xml').documentElement;
}

function makeBodyPr(inner: string): Element {
    const xml = `<a:bodyPr xmlns:a="${A_NS}">${inner}</a:bodyPr>`;
    return new DOMParser().parseFromString(xml, 'text/xml').documentElement;
}

function makeSpPr(inner: string): Element {
    const xml = `<a:spPr xmlns:a="${A_NS}">${inner}</a:spPr>`;
    return new DOMParser().parseFromString(xml, 'text/xml').documentElement;
}

describe('parseShape – textbox autoFit', () => {
    it('spAutoFit → autoFit = "shape"', () => {
        const node = makeWsp(`<a:bodyPr><a:spAutoFit/></a:bodyPr>`);
        const shape = parseShape(node, MOCK_OPTIONS, MOCK_CALLBACKS);
        expect(shape.props.textbox?.autoFit).toBe('shape');
    });

    it('normAutofit → autoFit = "normal"', () => {
        const node = makeWsp(`<a:bodyPr><a:normAutofit/></a:bodyPr>`);
        const shape = parseShape(node, MOCK_OPTIONS, MOCK_CALLBACKS);
        expect(shape.props.textbox?.autoFit).toBe('normal');
    });

    it('noAutofit → autoFit = "none"', () => {
        const node = makeWsp(`<a:bodyPr><a:noAutofit/></a:bodyPr>`);
        const shape = parseShape(node, MOCK_OPTIONS, MOCK_CALLBACKS);
        expect(shape.props.textbox?.autoFit).toBe('none');
    });

    it('no child element → autoFit = "none"', () => {
        const node = makeWsp(`<a:bodyPr/>`);
        const shape = parseShape(node, MOCK_OPTIONS, MOCK_CALLBACKS);
        expect(shape.props.textbox?.autoFit).toBe('none');
    });
});

describe('parseShape – textbox vertical anchor', () => {
    it('anchor="t" → verticalAnchor = "t"', () => {
        const node = makeWsp(`<a:bodyPr anchor="t"/>`);
        const shape = parseShape(node, MOCK_OPTIONS, MOCK_CALLBACKS);
        expect(shape.props.textbox?.verticalAnchor).toBe('t');
    });

    it('anchor="ctr" → verticalAnchor = "ctr"', () => {
        const node = makeWsp(`<a:bodyPr anchor="ctr"/>`);
        const shape = parseShape(node, MOCK_OPTIONS, MOCK_CALLBACKS);
        expect(shape.props.textbox?.verticalAnchor).toBe('ctr');
    });

    it('anchor="b" → verticalAnchor = "b"', () => {
        const node = makeWsp(`<a:bodyPr anchor="b"/>`);
        const shape = parseShape(node, MOCK_OPTIONS, MOCK_CALLBACKS);
        expect(shape.props.textbox?.verticalAnchor).toBe('b');
    });
});

describe('parseTransform2D – originalWidth / originalHeight', () => {
    it('stores original (pre-rotation) EMU dimensions as pt strings', () => {
        const node = makeSpPr(`<a:xfrm><a:ext cx="914400" cy="457200"/></a:xfrm>`);
        const shape: OpenXmlElement = { props: { is_transform: false, transform: {} }, cssStyle: {}, children: [] } as any;
        parseShapeProperties(node, shape, MOCK_OPTIONS);
        // 914400 EMU = 72pt, 457200 EMU = 36pt
        expect(parseFloat(shape.props.originalWidth)).toBeCloseTo(72, 1);
        expect(shape.props.originalWidth).toMatch(/pt$/);
        expect(parseFloat(shape.props.originalHeight)).toBeCloseTo(36, 1);
        expect(shape.props.originalHeight).toMatch(/pt$/);
    });

    it('stores original dimensions unchanged even when rotation expands bounding box', () => {
        // 45-degree rotation: bounding box grows, but originalWidth/Height should be pre-rotation
        const node = makeSpPr(`<a:xfrm rot="2700000"><a:ext cx="914400" cy="914400"/></a:xfrm>`);
        const shape: OpenXmlElement = { props: { is_transform: false, transform: {} }, cssStyle: {}, children: [] } as any;
        parseShapeProperties(node, shape, MOCK_OPTIONS);
        // originalWidth and originalHeight must be 72pt (pre-rotation cx/cy)
        expect(parseFloat(shape.props.originalWidth)).toBeCloseTo(72, 1);
        expect(parseFloat(shape.props.originalHeight)).toBeCloseTo(72, 1);
        // cssStyle width should be the rotated bounding box (larger)
        const cssWidth = parseFloat(shape.cssStyle['width'] ?? '0');
        const cssHeight = parseFloat(shape.cssStyle['height'] ?? '0');
        expect(cssWidth).toBeGreaterThanOrEqual(72);
        expect(cssHeight).toBeGreaterThanOrEqual(72);
    });
});
