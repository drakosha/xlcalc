// The keyboard's input source, for the steps that type on the keyboard (`entry` with `keys`, `fit` with typing, `ui`):
// `swift bench/coverage/input-source.swift` prints the current one's id, `swift bench/coverage/input-source.swift
// com.apple.keylayout.ABC` selects that one (it must be enabled in System Settings) and prints the id selected, so a run
// switches to ABC and back to what it printed before.
import Carbon

func currentId() -> String {
  let source = TISCopyCurrentKeyboardInputSource().takeRetainedValue()
  return Unmanaged<CFString>.fromOpaque(TISGetInputSourceProperty(source, kTISPropertyInputSourceID)).takeUnretainedValue() as String
}

let args = CommandLine.arguments
if args.count > 1 {
  let filter = [kTISPropertyInputSourceID as String: args[1]] as CFDictionary
  guard let list = TISCreateInputSourceList(filter, false)?.takeRetainedValue() as? [TISInputSource], let source = list.first else {
    FileHandle.standardError.write("no enabled input source \(args[1])\n".data(using: .utf8)!)
    exit(1)
  }
  TISSelectInputSource(source)
}
print(currentId())
