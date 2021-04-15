/**
 * PptxGenJS: Utility Methods
 */

import { EMU, REGEX_HEX_COLOR, DEF_FONT_COLOR, ONEPT, SchemeColor, SCHEME_COLORS } from './core-enums'
import { IChartOpts, ILayout, GlowOptions, ISlideLib, ShapeFill, Color, ShapeLine, ShapeGradient, GradientStops } from './core-interfaces'

/**
 * Convert string percentages to number relative to slide size
 * @param {number|string} size - numeric ("5.5") or percentage ("90%")
 * @param {'X' | 'Y'} xyDir - direction
 * @param {ILayout} layout - presentation layout
 * @returns {number} calculated size
 */
export function getSmartParseNumber(size: number | string, xyDir: 'X' | 'Y', layout: ILayout): number {
	// FIRST: Convert string numeric value if reqd
	if (typeof size === 'string' && !isNaN(Number(size))) size = Number(size)

	// CASE 1: Number in inches
	// Assume any number less than 100 is inches
	if (typeof size === 'number' && size < 100) return inch2Emu(size)

	// CASE 2: Number is already converted to something other than inches
	// Assume any number greater than 100 is not inches! Just return it (its EMU already i guess??)
	if (typeof size === 'number' && size >= 100) return size

	// CASE 3: Percentage (ex: '50%')
	if (typeof size === 'string' && size.indexOf('%') > -1) {
		if (xyDir && xyDir === 'X') return Math.round((parseFloat(size) / 100) * layout.width)
		if (xyDir && xyDir === 'Y') return Math.round((parseFloat(size) / 100) * layout.height)

		// Default: Assume width (x/cx)
		return Math.round((parseFloat(size) / 100) * layout.width)
	}

	// LAST: Default value
	return 0
}

/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
export function getUuid(uuidFormat: string): string {
	return uuidFormat.replace(/[xy]/g, function (c) {
		let r = (Math.random() * 16) | 0,
			v = c === 'x' ? r : (r & 0x3) | 0x8
		return v.toString(16)
	})
}

/**
 * TODO: What does this method do again??
 * shallow mix, returns new object
 */
export function getMix(o1: any | IChartOpts, o2: any | IChartOpts, etc?: any) {
	let objMix = {}
	for (let i = 0; i <= arguments.length; i++) {
		let oN = arguments[i]
		if (oN)
			Object.keys(oN).forEach(key => {
				objMix[key] = oN[key]
			})
	}
	return objMix
}

/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string} xml - XML string to encode
 * @returns {string} escaped XML
 */
export function encodeXmlEntities(xml: string): string {
	// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	if (typeof xml === 'undefined' || xml == null) return ''
	return xml.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

/**
 * Convert inches into EMU
 * @param {number|string} inches - as string or number
 * @returns {number} EMU value
 */
export function inch2Emu(inches: number | string): number {
	// FIRST: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
	// Any value over 100 damn sure isnt inches, must be EMU already, so just return it
	if (typeof inches === 'number' && inches > 100) return inches
	if (typeof inches === 'string') inches = Number(inches.replace(/in*/gi, ''))
	return Math.round(EMU * inches)
}

/**
 * Convert `pt` into points (using `ONEPT`)
 *
 * @param {number|string} pt
 * @returns {number} value in points (`ONEPT`)
 */
export function valToPts(pt: number | string): number {
	let points = Number(pt) || 0
	return isNaN(points) ? 0 : Math.round(points * ONEPT)
}

/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 *
 * @param {number} d - degrees
 * @returns {number} rot - value
 */
export function convertRotationDegrees(d: number): number {
	d = d || 0
	return (d > 360 ? d - 360 : d) * 60000
}

/**
 * Converts component value to hex value
 * @param {number} c - component color
 * @returns {string} hex string
 */
export function componentToHex(c: number): string {
	let hex = c.toString(16)
	return hex.length === 1 ? '0' + hex : hex
}

/**
 * Converts RGB colors from css selectors to Hex for Presentation colors
 * @param {number} r - red value
 * @param {number} g - green value
 * @param {number} b - blue value
 * @returns {string} XML string
 */
export function rgbToHex(r: number, g: number, b: number): string {
	return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase()
}

/**
 * Create either a `a:schemeClr` - (scheme color) or `a:srgbClr` (hexa representation).
 * @param {string|SCHEME_COLORS} colorStr - hexa representation (eg. "FFFF00") or a scheme color constant (eg. pptx.SchemeColor.ACCENT1)
 * @param {string} innerElements - additional elements that adjust the color and are enclosed by the color element
 * @returns {string} XML string
 */
export function createColorElement(colorStr: string | SCHEME_COLORS, innerElements?: string): string {
	let colorVal = (colorStr || '').replace('#', '')
	let isHexaRgb = REGEX_HEX_COLOR.test(colorVal)

	if (
		!isHexaRgb &&
		colorVal !== SchemeColor.background1 &&
		colorVal !== SchemeColor.background2 &&
		colorVal !== SchemeColor.text1 &&
		colorVal !== SchemeColor.text2 &&
		colorVal !== SchemeColor.accent1 &&
		colorVal !== SchemeColor.accent2 &&
		colorVal !== SchemeColor.accent3 &&
		colorVal !== SchemeColor.accent4 &&
		colorVal !== SchemeColor.accent5 &&
		colorVal !== SchemeColor.accent6
	) {
		console.warn(`"${colorVal}" is not a valid scheme color or hexa RGB! "${DEF_FONT_COLOR}" is used as a fallback. Pass 6-digit RGB or 'pptx.SchemeColor' values`)
		colorVal = DEF_FONT_COLOR
	}

	let tagName = isHexaRgb ? 'srgbClr' : 'schemeClr'
	let colorAttr = 'val="' + (isHexaRgb ? colorVal.toUpperCase() : colorVal) + '"'

	return innerElements ? `<a:${tagName} ${colorAttr}>${innerElements}</a:${tagName}>` : `<a:${tagName} ${colorAttr}/>`
}

/**
 * Creates `a:glow` element
 * @param {GlowOptions} options glow properties
 * @param {GlowOptions} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 *	{ size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
export function createGlowElement(options: GlowOptions, defaults: GlowOptions): string {
	let strXml = '',
		opts = getMix(defaults, options),
		size = opts['size'] * ONEPT,
		color = opts['color'],
		opacity = opts['opacity'] * 100000

	strXml += `<a:glow rad="${size}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += `</a:glow>`

	return strXml
}

/**
 * Create color selection
 * @param {shapeFill} ShapeFill - options
 * @param {string} backColor - color string
 * @returns {string} XML string
 */
export function genXmlColorSelection(shapeFill: Color | ShapeFill | ShapeLine | ShapeGradient, backColor?: string): string {
	let colorVal = ''
	let fillType = 'solid'
	let internalElements = ''
	let outText = ''
	let shapeGradient = shapeFill as ShapeGradient
	let newGradientOpts:ShapeGradient = {}
	let defaultGradStops = [
		{color:SchemeColor["accent1"], position:0, transparency:0, brightness:0 },
		{color:SchemeColor["accent1"], position:50, transparency:0, brightness:50 },
		{color:SchemeColor["accent1"], position:100, transparency:0, brightness:100 }
	]

	if (backColor && typeof backColor === 'string') {
		outText += `<p:bg><p:bgPr>${genXmlColorSelection(backColor.replace('#', ''))}<a:effectLst/></p:bgPr></p:bg>`
	} else if (backColor && typeof backColor === 'object') {
		outText += `<p:bg><p:bgPr>${genXmlColorSelection(backColor)}"<a:effectLst/></p:bgPr></p:bg>`
	}

	if (shapeFill) {
		if (typeof shapeFill === 'string') colorVal = shapeFill
		else {
			if (shapeFill.type) fillType = shapeFill.type
			if (shapeFill.color) colorVal = shapeFill.color
			if (shapeFill.alpha) internalElements += `<a:alpha val="${100 - shapeFill.alpha}000"/>` // @deprecated v3.3.0
			if (shapeFill.transparency) internalElements += `<a:alpha val="${100 - shapeFill.transparency}000"/>`
			if (fillType === 'gradient') {
				newGradientOpts.gradientType = shapeGradient.gradientType || 'linear',
				newGradientOpts.rotateWithShape = shapeGradient.rotateWithShape || 1,
				newGradientOpts.linearAngle = null, // Supercedes gradientDirection for linear gradients if supplied
				newGradientOpts.gradientDirection = shapeGradient.gradientDirection || null,
				newGradientOpts.pathL = shapeGradient.pathL || 50,
				newGradientOpts.pathT = shapeGradient.pathT || 50,
				newGradientOpts. gradStops = shapeGradient.gradStops || defaultGradStops
				if (newGradientOpts.gradientType == 'linear') {
					if ((shapeGradient.linearAngle) && (shapeGradient.linearAngle >=0) && (shapeGradient.linearAngle <360)) { 
						newGradientOpts.linearAngle = shapeGradient.linearAngle;
					} else if (newGradientOpts.gradientDirection === 'lr') { newGradientOpts.linearAngle = 0;
					} else if (newGradientOpts.gradientDirection === 'tlbr') { newGradientOpts.linearAngle = 45;
					} else if (newGradientOpts.gradientDirection === 'tb') { newGradientOpts.linearAngle = 90;
					} else if (newGradientOpts.gradientDirection === 'trbl') { newGradientOpts.linearAngle = 135;
					} else if (newGradientOpts.gradientDirection === 'rl') { newGradientOpts.linearAngle = 180;
					} else if (newGradientOpts.gradientDirection === 'brtl') { newGradientOpts.linearAngle = 225;
					} else if (newGradientOpts.gradientDirection === 'bt') { newGradientOpts.linearAngle = 270;
					} else if (newGradientOpts.gradientDirection === 'bltr') { newGradientOpts.linearAngle = 315;
					} else {
						// Fail Safe in case a Radial or Rectangular directional value was entered
						newGradientOpts.linearAngle = 45
					}
				} else if ((newGradientOpts.gradientType === 'radial') || (newGradientOpts.gradientType === 'rect')) {
					if (!(newGradientOpts.gradientDirection === 'ftl') && !(newGradientOpts.gradientDirection === 'ftr') && !(newGradientOpts.gradientDirection === 'fbl') && !(newGradientOpts.gradientDirection === 'fbr') && !(newGradientOpts.gradientDirection === 'c')) {
						newGradientOpts.gradientDirection = 'ftl';
					}
				} else if (newGradientOpts.gradientType == 'path') {
					if ((newGradientOpts.pathL <0) || (newGradientOpts.pathL >100)) {
						newGradientOpts.pathL = 50
					}
					newGradientOpts.pathR = 100 - newGradientOpts.pathL
					if ((newGradientOpts.pathT <0) || (newGradientOpts.pathT >100)) {
						newGradientOpts.pathT = 50
					}
					newGradientOpts.pathB = 100 - newGradientOpts.pathT
				}
				
				if (newGradientOpts.gradStops.length < 2 && newGradientOpts.gradStops.length > 10) {
					newGradientOpts.gradStops = defaultGradStops;
				}
				newGradientOpts.gradStops.forEach(function(stopPoint, i){
					stopPoint.position = ((stopPoint.position) && (Number(stopPoint.position) >= 0) && (Number(stopPoint.position) <= 100)) ? Number(stopPoint.position) : i*10
					stopPoint.brightness = ((stopPoint.brightness) && (Number(stopPoint.brightness) >= -100) && (Number(stopPoint.brightness) <= 100)) ? Number(stopPoint.brightness) : 0 
					stopPoint.transparency = ((stopPoint.transparency) && (Number(stopPoint.transparency) >= 0) && (Number(stopPoint.transparency) <= 100)) ? Number(stopPoint.transparency) : 0
				})
			}
		}
			switch (fillType) {
				case 'solid':
					outText += `<a:solidFill>${createColorElement(colorVal, internalElements)}</a:solidFill>`
					break;
				case 'gradient':	
					outText += `<a:gradFill flip="none" rotWithShape="${newGradientOpts.rotateWithShape}">`;
					outText += `<a:gsLst>`;
					newGradientOpts.gradStops.forEach(function(stopPoint){
						var stopPointInternalElements = '';
						outText += `<a:gs pos="${(stopPoint.position * 1000)}">`;
						if (stopPoint.brightness < 0) {
							stopPointInternalElements += `<a:lumMod val="${((100 - stopPoint.brightness * -1) * 1000)}"/>`;
						} else if (stopPoint.brightness > 0) {
							stopPointInternalElements += `<a:lumMod val=" + ${((100 - stopPoint.brightness) * 1000)}"/>`;
							stopPointInternalElements += `<a:lumOff val=" + ${(stopPoint.brightness * 1000)}"/>`;
						}
						if (stopPoint.transparency > 0) {
							stopPointInternalElements += `<a:alpha val="${((100 - stopPoint.transparency) * 1000)}"/>`;
						}
						outText += createColorElement(stopPoint.color, stopPointInternalElements);
						outText += `</a:gs>`;
					})
					outText += `</a:gsLst>`;
	
					switch(newGradientOpts.gradientType) {
						case 'linear':
							outText += `<a:lin ang="${(newGradientOpts.linearAngle * 60000)}" scaled="1"/>`;
							outText += `<a:tileRect/>`
							break;
						case 'radial':
						case 'rect':
							outText += `<a:path path="${(newGradientOpts.gradientType == "radial" ? "circle" : "rect")}">`;
							if (newGradientOpts.gradientDirection === 'ftl') { 
								outText += '<a:fillToRect r="100000" b="100000"/></a:path><a:tileRect l="-100000" t="-100000"/>'
							} else if (newGradientOpts.gradientDirection === 'ftr') {
								outText += '<a:fillToRect l="100000" b="100000"/></a:path><a:tileRect t="-100000" r="-100000"/>'
							} else if (newGradientOpts.gradientDirection === 'fbl') {
								outText += '<a:fillToRect t="100000" r="100000"/></a:path><a:tileRect l="-100000" b="-100000"/>'
							} else if (newGradientOpts.gradientDirection === 'fbr') { 
								outText += '<a:fillToRect l="100000" t="100000"/></a:path><a:tileRect r="-100000" b="-100000"/>'
							} else if (newGradientOpts.gradientDirection === 'c') { 
								outText += '<a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path><a:tileRect/>'
							} else {
								// Fail Safe in case a Radial or Rectangular directional value was entered
								outText += '<a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path><a:tileRect/>'
							}
							break;
						case 'path':
							outText += `<a:path path="shape">`;
							outText += `<a:fillToRect l="${(newGradientOpts.pathL * 1000)}" t="${(newGradientOpts.pathT * 1000)}" r="${(newGradientOpts.pathR * 1000)}" b="${(newGradientOpts.pathB * 1000)}"/>`;
							outText += `</a:path>`;
							outText += `<a:tileRect/>`;
							break;
					}
					outText += `</a:gradFill>`;
					break;
			default:
				outText += '' // @note need a statement as having only "break" is removed by rollup, then tiggers "no-default" js-linter
				break
			}
		//}
	}

	return outText
}

/**
 * Get a new rel ID (rId) for charts, media, etc.
 * @param {ISlideLib} target - the slide to use
 * @returns {number} count of all current rels plus 1 for the caller to use as its "rId"
 */
export function getNewRelId(target: ISlideLib): number {
	return target.rels.length + target.relsChart.length + target.relsMedia.length + 1
}
