import Foundation
import Vision
import AppKit

let args = Array(CommandLine.arguments.dropFirst())
if args.isEmpty { FileHandle.standardError.write("usage: ocr <image...>\n".data(using:.utf8)!); exit(2) }
for p in args {
    guard let img = NSImage(contentsOfFile: p),
          let cg = img.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
        print("=== \(p) :: IMAGE_UNREADABLE")
        continue
    }
    print("=== FILE \(p) size=\(cg.width)x\(cg.height)")
    let req = VNRecognizeTextRequest()
    req.recognitionLevel = .accurate
    req.usesLanguageCorrection = false
    let handler = VNImageRequestHandler(cgImage: cg, options: [:])
    do {
        try handler.perform([req])
        guard let obs = req.results, !obs.isEmpty else { print("  (no text found)"); continue }
        let lines = obs.compactMap { $0.topCandidates(1).first?.string }
        for l in lines { print("  " + l) }
    } catch {
        print("  OCR_ERROR \(error)")
    }
}
