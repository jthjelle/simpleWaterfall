# Waterfall Chart Visual Walkthrough

This document outlines the features and improvements implemented in the Waterfall Chart custom visual for Power BI.

## Completed Features

### 1. Metric Expansion
- Added support for **Budget** and **Estimate** fields.
- Users can now select which metric to use for the **Start Value** and **End Value** of the waterfall.
    - Example: Start with `Budget`, end with `Estimate`.
- Updated data processing to handle these new roles.

### 2. Fixed Y-Axis Range
- Added **Min** and **Max** settings to the Y-Axis configuration.
- Users can now lock the axis scale to a specific range (e.g., 0 to 100), preventing auto-scaling shifts when filtering data.

### 3. Orientation (Vertical Layout)
- Added an **Orientation** toggle in the "Layout Settings" card.
- Supports **Horizontal** (standard) and **Vertical** layouts.
    - **Horizontal**: Categories on Y-axis, Values on X-axis.
    - **Vertical**: Categories on X-axis, Values on Y-axis.
- Axes, bars, connectors, and labels automatically adjust to the selected orientation.

### 4. Advanced Label Styling
- **Font Family**: Added logic to respect the standard font family setting for data labels.
- **Background Badges**:
    - Introduced a "Show Background" toggle for data labels.
    - Configurable **Background Color** and **Transparency**.
    - Renders a rounded rectangle behind each label to improve readability against chart bars.

### 5. Axis Break Improvements
- Implemented visual indicators (masking) for axis breaks on total bars.
- Fixed logic to safely handle axis breaks in vertical orientation.

## Verification
- **Build**: The project builds successfully with `pbiviz package`.
- **Functionality**:
    - Verified data extraction for all new metrics.
    - Confirmed toggling between Horizontal and Vertical layouts updates the rendering correctly.
    - Confirmed Label Badges appear with correct color/opacity behind text.
    - Confirmed Axis Min/Max settings override auto-scaling.

## Next Steps
- Deploy and test in Power BI Service.
- Gather user feedback on the new styling options.
