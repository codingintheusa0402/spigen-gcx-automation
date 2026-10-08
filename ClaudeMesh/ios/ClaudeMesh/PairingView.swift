import SwiftUI
import AVFoundation

struct PairingView: View {
    @EnvironmentObject var pairing: Pairing
    @State private var scanning = false
    @State private var pasted = ""
    @State private var error: String?

    var body: some View {
        ZStack {
            LinearGradient(colors: [Color(red: 0.10, green: 0.09, blue: 0.31), Color(red: 0.02, green: 0.02, blue: 0.10)], startPoint: .top, endPoint: .bottom).ignoresSafeArea()
            VStack(spacing: 22) {
                Image("AppIconImage").resizable().frame(width: 84, height: 84).clipShape(RoundedRectangle(cornerRadius: 20))
                    .shadow(color: .blue.opacity(0.5), radius: 18)
                Text("G C X   M E S H").font(.system(size: 17, weight: .light)).foregroundStyle(.white)
                Text("On your Mac open GCX Mesh, click 📱, turn on phone access, then scan the QR code.\nUse the Tailscale code to connect from anywhere.")
                    .font(.footnote).multilineTextAlignment(.center).foregroundStyle(.white.opacity(0.65)).padding(.horizontal, 28)
                Button { scanning = true } label: {
                    Label("Scan QR code", systemImage: "qrcode.viewfinder").font(.headline).padding(.horizontal, 26).padding(.vertical, 13)
                        .background(Capsule().fill(Color(red: 0.435, green: 0.608, blue: 1)))
                        .foregroundStyle(Color(red: 0.03, green: 0.06, blue: 0.23))
                }
                VStack(spacing: 8) {
                    TextField("…or paste the pairing link", text: $pasted).textInputAutocapitalization(.never).autocorrectionDisabled()
                        .padding(12).background(RoundedRectangle(cornerRadius: 14).fill(.white.opacity(0.08))).foregroundStyle(.white)
                    Button("Connect") { if !pairing.add(pasted) { error = "That isn't a GCX Mesh pairing link." } }.disabled(pasted.isEmpty)
                }.padding(.horizontal, 28)
                if let error { Text(error).font(.footnote).foregroundStyle(.red.opacity(0.85)) }
            }
        }
        .sheet(isPresented: $scanning) {
            QRScanner { code in
                scanning = false
                if !pairing.add(code) { error = "That QR code isn't a GCX Mesh pairing code." }
            }
            .ignoresSafeArea()
        }
    }
}

/// Camera QR scanner (AVFoundation).
struct QRScanner: UIViewControllerRepresentable {
    var onCode: (String) -> Void
    func makeUIViewController(context: Context) -> ScannerVC { let vc = ScannerVC(); vc.onCode = onCode; return vc }
    func updateUIViewController(_ vc: ScannerVC, context: Context) {}

    final class ScannerVC: UIViewController, AVCaptureMetadataOutputObjectsDelegate {
        var onCode: ((String) -> Void)?
        private let session = AVCaptureSession()
        private var done = false
        override func viewDidLoad() {
            super.viewDidLoad(); view.backgroundColor = .black
            guard let cam = AVCaptureDevice.default(for: .video), let input = try? AVCaptureDeviceInput(device: cam), session.canAddInput(input) else {
                let l = UILabel(); l.text = "Camera unavailable — paste the link instead."; l.textColor = .white; l.textAlignment = .center
                l.frame = view.bounds; l.autoresizingMask = [.flexibleWidth, .flexibleHeight]; view.addSubview(l); return
            }
            session.addInput(input)
            let out = AVCaptureMetadataOutput()
            if session.canAddOutput(out) { session.addOutput(out); out.setMetadataObjectsDelegate(self, queue: .main); out.metadataObjectTypes = [.qr] }
            let layer = AVCaptureVideoPreviewLayer(session: session); layer.frame = view.bounds; layer.videoGravity = .resizeAspectFill
            view.layer.addSublayer(layer)
            DispatchQueue.global(qos: .userInitiated).async { self.session.startRunning() }
        }
        override func viewWillDisappear(_ animated: Bool) { super.viewWillDisappear(animated); session.stopRunning() }
        func metadataOutput(_ output: AVCaptureMetadataOutput, didOutput objects: [AVMetadataObject], from connection: AVCaptureConnection) {
            guard !done, let s = (objects.first as? AVMetadataMachineReadableCodeObject)?.stringValue else { return }
            done = true; UINotificationFeedbackGenerator().notificationOccurred(.success); onCode?(s)
        }
    }
}
