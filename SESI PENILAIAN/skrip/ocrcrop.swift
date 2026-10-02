import Foundation
import Vision
import AppKit

// Usage: ocrcrop <image> <gridCols> <gridRows>
let args = Array(CommandLine.arguments.dropFirst())
guard args.count >= 3, let img = NSImage(contentsOfFile: args[0]),
      let cg = img.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
    print("usage error"); exit(2)
}
let cols = Int(args[1]) ?? 2, rows = Int(args[2]) ?? 2
let W = cg.width, H = cg.height
print("=== \(args[0]) \(W)x\(H) grid \(cols)x\(rows)")
for ry in 0..<rows {
    for rx in 0..<cols {
        let x = W * rx / cols, y = H * ry / rows
        let w = W / cols, h = H / rows
        guard let sub = cg.cropping(to: CGRect(x: x, y: y, width: w, height: h)) else { continue }
        // upscale sub 2x
        let rep = NSBitmapImageRep(cgImage: sub)
        let target = NSImage(size: NSSize(width: w * 2, height: h * 2))
        target.lockFocus()
        NSGraphicsContext.current?.imageInterpolation = .high
        rep.draw(in: NSRect(x: 0, y: 0, width: w * 2, height: h * 2))
        target.unlockFocus()
        guard let up = target.cgImage(forProposedRect: nil, context: nil, hints: nil) else { continue }
        let req = VNRecognizeTextRequest()
        req.recognitionLevel = .accurate
        req.usesLanguageCorrection = false
        let handler = VNImageRequestHandler(cgImage: up, options: [:])
        try? handler.perform([req])
        let lines = (req.results ?? []).compactMap { $0.topCandidates(1).first?.string }
        print("  [tile r\(ry)c\(rx)]")
        for l in lines { print("      " + l) }
    }
}
