# Sheets - Modern Spreadsheet Application

A clean, modern web-based spreadsheet application that looks and operates similar to Google Sheets and Airtable. Built with vanilla HTML, CSS, and JavaScript.

## Features

### 🎯 **Core Functionality**
- **Interactive Cells**: Click any cell to select it, double-click to edit
- **Keyboard Navigation**: Use arrow keys, Tab, and Enter to navigate between cells
- **Real-time Editing**: Type directly in cells to update content
- **Data Persistence**: All data is stored in memory during the session

### 🎨 **User Interface**
- **Clean Design**: Minimalist white background with subtle gray borders
- **Responsive Layout**: Works on desktop and mobile devices
- **Modern Icons**: Font Awesome icons for a professional look
- **Smooth Animations**: Hover effects and transitions throughout

### 🔍 **Search & Navigation**
- **Global Search**: Search across all cells with real-time highlighting
- **Column Selection**: Click column headers to select entire columns
- **Row Selection**: Click row numbers to select entire rows
- **Cell Selection**: Click individual cells for precise selection

### ⚡ **Interactive Elements**
- **Action Buttons**: "Run all" and "Add column" functionality
- **Floating Actions**: Bottom-right action buttons for quick access
- **Context Menu**: Right-click on cells for copy, paste, cut, delete, and format options
- **Notifications**: Toast notifications for user feedback

### 📱 **Responsive Design**
- **Mobile Friendly**: Optimized for touch devices
- **Flexible Layout**: Adapts to different screen sizes
- **Touch Support**: Works seamlessly on tablets and phones

## Getting Started

### Prerequisites
- Modern web browser (Chrome, Firefox, Safari, Edge)
- No additional software installation required

### Installation
1. Download or clone the repository
2. Open `index.html` in your web browser
3. Start using the spreadsheet immediately!

### File Structure
```
Sheets/
├── index.html          # Main HTML structure
├── styles.css          # CSS styling and layout
├── script.js           # JavaScript functionality
└── README.md           # This documentation
```

## Usage Guide

### Basic Navigation
- **Click** any cell to select it
- **Double-click** to edit cell content
- **Arrow keys** to move between cells
- **Tab** to move right, **Shift+Tab** to move left
- **Enter** to move down, **Shift+Enter** to move up

### Cell Operations
- **Type** directly in cells to add content
- **Right-click** for context menu options
- **Copy/Paste** using context menu or keyboard shortcuts
- **Delete** content using context menu

### Advanced Features
- **Search**: Use the search bar to find specific content
- **Add Columns**: Click "Add column" button to expand the spreadsheet
- **Run All**: Click "Run all" to simulate batch operations
- **Floating Actions**: Use bottom action buttons for quick access

### Keyboard Shortcuts
| Key | Action |
|-----|--------|
| `↑` | Move up |
| `↓` | Move down |
| `←` | Move left |
| `→` | Move right |
| `Tab` | Move right |
| `Shift+Tab` | Move left |
| `Enter` | Move down |
| `Shift+Enter` | Move up |

## Customization

### Adding More Rows/Columns
The spreadsheet is configured for 18 rows and 8 columns by default. To modify:
1. Open `script.js`
2. Change the `rows` and `columns` values in the constructor
3. Refresh the page

### Styling Changes
- Modify `styles.css` to change colors, fonts, and layout
- Update CSS variables for consistent theming
- Adjust responsive breakpoints for different screen sizes

### Functionality Extensions
- Add new features in `script.js`
- Implement data export/import functionality
- Add formula support for calculations
- Integrate with external data sources

## Browser Compatibility

- ✅ Chrome 80+
- ✅ Firefox 75+
- ✅ Safari 13+
- ✅ Edge 80+
- ✅ Mobile browsers (iOS Safari, Chrome Mobile)

## Performance Features

- **Efficient Rendering**: Only renders visible cells
- **Smooth Scrolling**: Optimized for large datasets
- **Memory Management**: Efficient data storage and retrieval
- **Event Delegation**: Minimal event listeners for better performance

## Future Enhancements

- [ ] Data export to CSV/Excel
- [ ] Formula support and calculations
- [ ] Cell formatting options
- [ ] Collaborative editing
- [ ] Data validation rules
- [ ] Chart and graph integration
- [ ] Undo/redo functionality
- [ ] Auto-save to local storage

## Contributing

Feel free to contribute to this project by:
- Reporting bugs
- Suggesting new features
- Submitting pull requests
- Improving documentation

## License

This project is open source and available under the MIT License.

## Support

For questions or support, please open an issue in the repository or contact the development team.

---

**Built with ❤️ using vanilla web technologies**
