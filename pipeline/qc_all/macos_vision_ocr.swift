import Foundation
import Vision
import ImageIO

// macOS Vision OCR helper. Output: one NDJSON object per input image.
func jsonLine(_ object: [String: Any]) {
    do {
        let data = try JSONSerialization.data(withJSONObject: object, options: [])
        FileHandle.standardOutput.write(data)
        FileHandle.standardOutput.write(Data([0x0A]))
    } catch {
        let fallback = "{\"status\":\"json_error\",\"error\":\"\(error.localizedDescription)\"}\n"
        FileHandle.standardOutput.write(fallback.data(using: .utf8)!)
    }
}

let imagePaths = Array(CommandLine.arguments.dropFirst())
if imagePaths.isEmpty {
    FileHandle.standardError.write("Usage: macos_vision_ocr.swift IMAGE [IMAGE ...]\n".data(using: .utf8)!)
    exit(2)
}

for (index, path) in imagePaths.enumerated() {
    autoreleasepool {
        let url = URL(fileURLWithPath: path)
        guard let source = CGImageSourceCreateWithURL(url as CFURL, nil),
              let image = CGImageSourceCreateImageAtIndex(source, 0, nil) else {
            jsonLine(["source": path, "page_index": index + 1, "status": "image_open_failed", "blocks": []])
            return
        }

        let request = VNRecognizeTextRequest()
        request.recognitionLevel = .accurate
        request.usesLanguageCorrection = true
        request.recognitionLanguages = ["zh-Hans", "en-US"]

        let handler = VNImageRequestHandler(cgImage: image, options: [:])
        do {
            try handler.perform([request])
            let observations = (request.results ?? []).sorted { lhs, rhs in
                if abs(lhs.boundingBox.midY - rhs.boundingBox.midY) > 0.006 {
                    return lhs.boundingBox.midY > rhs.boundingBox.midY
                }
                return lhs.boundingBox.minX < rhs.boundingBox.minX
            }
            let blocks: [[String: Any]] = observations.compactMap { observation in
                guard let candidate = observation.topCandidates(1).first else { return nil }
                let box = observation.boundingBox
                return [
                    "text": candidate.string,
                    "confidence": Double(candidate.confidence),
                    "x": box.minX, "y": box.minY,
                    "width": box.width, "height": box.height,
                    "mid_x": box.midX, "mid_y": box.midY,
                ]
            }
            jsonLine([
                "source": path, "page_index": index + 1, "status": "ok",
                "image_width": image.width, "image_height": image.height,
                "coordinate_origin": "bottom_left_normalized",
                "sort_order": "top_to_bottom_then_left_to_right",
                "languages": ["zh-Hans", "en-US"],
                "block_count": blocks.count, "blocks": blocks,
            ])
        } catch {
            jsonLine([
                "source": path, "page_index": index + 1, "status": "ocr_failed",
                "error": error.localizedDescription, "blocks": [],
            ])
        }
    }
}
