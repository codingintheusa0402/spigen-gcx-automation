import AppKit

// Renders the 1024x1024 app icon: Nova-blue gradient, white vault/column glyph, small "SIM" tag.
let size = CGSize(width: 1024, height: 1024)
let image = NSImage(size: size)
image.lockFocus()
let ctx = NSGraphicsContext.current!.cgContext

let top = NSColor(red: 0.20, green: 0.47, blue: 1.00, alpha: 1)
let bottom = NSColor(red: 0.36, green: 0.22, blue: 0.95, alpha: 1)
let gradient = NSGradient(starting: top, ending: bottom)!
gradient.draw(in: NSRect(origin: .zero, size: size), angle: -60)

// soft highlight circle
ctx.saveGState()
ctx.setFillColor(NSColor.white.withAlphaComponent(0.10).cgColor)
ctx.fillEllipse(in: CGRect(x: 560, y: 560, width: 720, height: 720))
ctx.restoreGState()

let config = NSImage.SymbolConfiguration(pointSize: 560, weight: .semibold)
let symbol = NSImage(systemSymbolName: "building.columns.fill", accessibilityDescription: nil)!
    .withSymbolConfiguration(config)!
let tinted = NSImage(size: symbol.size, flipped: false) { rect in
    symbol.draw(in: rect)
    NSColor.white.set()
    rect.fill(using: .sourceAtop)
    return true
}
let s = tinted.size
let origin = CGPoint(x: (size.width - s.width) / 2, y: (size.height - s.height) / 2 + 40)
tinted.draw(in: NSRect(origin: origin, size: s))

// SIM tag
let tag = NSBezierPath(roundedRect: NSRect(x: 352, y: 118, width: 320, height: 112), xRadius: 56, yRadius: 56)
NSColor(red: 1.0, green: 0.62, blue: 0.10, alpha: 1).setFill()
tag.fill()
let para = NSMutableParagraphStyle(); para.alignment = .center
let attrs: [NSAttributedString.Key: Any] = [
    .font: NSFont.systemFont(ofSize: 76, weight: .black),
    .foregroundColor: NSColor.white,
    .paragraphStyle: para,
]
NSAttributedString(string: "SIM", attributes: attrs).draw(in: NSRect(x: 352, y: 128, width: 320, height: 100))

image.unlockFocus()
let tiff = image.tiffRepresentation!
let rep = NSBitmapImageRep(data: tiff)!
rep.size = size
let png = rep.representation(using: .png, properties: [:])!
let out = CommandLine.arguments.count > 1 ? CommandLine.arguments[1] : "icon-1024.png"
try! png.write(to: URL(fileURLWithPath: out))
print("wrote \(out) \(rep.pixelsWide)x\(rep.pixelsHigh)")
