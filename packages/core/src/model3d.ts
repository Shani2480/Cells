import type { Model3D } from "./shapes"

/**
 * Office's description of a 3D model's view (`am3d:model3d`), as Excel writes it for a model it inserts: the model
 * scaled to a metre on its longest side and centred, a 45° camera just clear of it, Office's three lights.
 * Lengths are 36,000,000 to the metre and angles 60,000 to the degree.
 */
export function model3dXml(fit: { meterPerUnit: number; preTrans: [number, number, number]; cameraZ: number }) {
  const len = (m: number) => Math.round(m * 36_000_000)
  // meterPerModelUnit as a fraction over a million, as Excel writes it.
  const n = Math.max(1, Math.round(fit.meterPerUnit * 1_000_000))
  const [dx, dy, dz] = fit.preTrans.map(len)
  return (
    `<am3d:model3d xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="">` +
    `<am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></am3d:spPr>` +
    `<am3d:camera><am3d:pos x="0" y="0" z="${len(fit.cameraZ)}"/><am3d:up dx="0" dy="36000000" dz="0"/><am3d:lookAt x="0" y="0" z="0"/><am3d:perspective fov="2700000"/></am3d:camera>` +
    `<am3d:trans><am3d:meterPerModelUnit n="${n}" d="1000000"/><am3d:preTrans dx="${dx}" dy="${dy}" dz="${dz}"/>` +
    `<am3d:scale><am3d:sx n="1000000" d="1000000"/><am3d:sy n="1000000" d="1000000"/><am3d:sz n="1000000" d="1000000"/></am3d:scale>` +
    `<am3d:rot/><am3d:postTrans dx="0" dy="0" dz="0"/></am3d:trans>` +
    `<am3d:raster rName="Office3DRenderer" rVer="16.0.8326"><am3d:blip r:embed=""/></am3d:raster>` +
    `<am3d:objViewport viewportSz="0"/>` +
    `<am3d:ambientLight><am3d:clr><a:scrgbClr r="50000" g="50000" b="50000"/></am3d:clr><am3d:illuminance n="500000" d="1000000"/></am3d:ambientLight>` +
    `<am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="100000" g="75000" b="50000"/></am3d:clr><am3d:intensity n="9765625" d="1000000"/><am3d:pos x="21959998" y="70920001" z="16344003"/></am3d:ptLight>` +
    `<am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="40000" g="60000" b="95000"/></am3d:clr><am3d:intensity n="12250000" d="1000000"/><am3d:pos x="-37964106" y="51130435" z="57631972"/></am3d:ptLight>` +
    `<am3d:ptLight rad="0"><am3d:clr><a:scrgbClr r="86837" g="72700" b="100000"/></am3d:clr><am3d:intensity n="3125000" d="1000000"/><am3d:pos x="-37739122" y="58056624" z="-34769649"/></am3d:ptLight>` +
    `</am3d:model3d>`
  )
}

/** A model's rotation as Office stores it: degrees kept between 0 and 360. */
export function normalRotation(rot: Model3D["rot"]): Model3D["rot"] {
  const n = (a: number) => {
    const v = ((a % 360) + 360) % 360
    return Math.round(v * 100) / 100 === 360 ? 0 : Math.round(v * 100) / 100
  }
  return { x: n(rot.x), y: n(rot.y), z: n(rot.z) }
}
