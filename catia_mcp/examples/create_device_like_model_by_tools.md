# Example: create a phone-like model using generic MCP tools

This MCP server does not contain an iPhone-specific tool.

A Codex task can create a phone-like model by composing generic tools:

1. catia_create_rounded_rect_slab
   - name: "DeviceBody"
   - width: 77.6
   - height: 163.0
   - thickness: 8.3
   - corner_radius: 18.0

2. catia_create_sketch
   - plane: "xy"
   - sketch_name: "CameraIslandSketch"
   - offset: 8.3

3. catia_sketch_rounded_rectangle
   - width: 38
   - height: 38
   - radius: 8
   - center_x: -18
   - center_y: 50

4. catia_close_sketch

5. catia_pad
   - height: 1.8
   - sketch_name: "CameraIslandSketch"
   - feature_name: "CameraIsland"

6. catia_circular_pad
   - radius: 6.75
   - height: 1.2
   - center_x: -26
   - center_y: 58
   - plane: "xy"
   - offset: 10.1
   - feature_name: "Lens1"

Repeat catia_circular_pad for other lenses.