# RVCE Counsellor Data Combiner

<div align="center">
  <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 30px;">
  <a href="https://rvce.edu.in" target="_blank" rel="noopener noreferrer" style="margin-right: 20px;">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://github.com/overclocked-2124/RVCE-Coding-Bootkit/blob/main/gitAssets/RVCE_Logo_With_Text.png">
      <img src="https://github.com/overclocked-2124/RVCE-Coding-Bootkit/blob/main/gitAssets/RVCE_Logo_With_Text_Black.png" alt="RVCE Text Logo" height="80">
    </picture>
  </a>
  <a href="https://www.linkedin.com/company/coding-club-rvce/posts/?feedView=all" target="_blank" rel="noopener noreferrer" style="margin-left: 20px;">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://github.com/overclocked-2124/RVCE-Coding-Bootkit/blob/main/gitAssets/CCLogo_BG_Removed.png">
      <img src="https://github.com/overclocked-2124/RVCE-Coding-Bootkit/blob/main/gitAssets/CCLogo_BG_Removed-Black.png" alt="Coding Club Logo" height="80">
    </picture>
  </a>
</div>

  <h3>Professional Excel Processing Tool</h3>
  <p>Efficiently combine student-counsellor data from multiple Excel files with advanced formatting and download capabilities</p>
  
[![Live Demo](https://img.shields.io/badge/Live-Demo-2563eb?style=for-the-badge&logo=vercel)](https://combined-counsellor-data-rvce.vercel.app)
[![React](https://img.shields.io/badge/React-18-61DAFB?style=for-the-badge&logo=react)](https://reactjs.org/)
[![ExcelJS](https://img.shields.io/badge/ExcelJS-4.4.0-217346?style=for-the-badge&logo=microsoftexcel)](https://github.com/exceljs/exceljs)
[![Vite](https://img.shields.io/badge/Vite-5.0-646CFF?style=for-the-badge&logo=vite)](https://vitejs.dev/)
[![TailwindCSS](https://img.shields.io/badge/Tailwind-CSS-06B6D4?style=for-the-badge&logo=tailwindcss)](https://tailwindcss.com/)
[![Lucide](https://img.shields.io/badge/Lucide-React-F56565?style=for-the-badge&logo=lucide)](https://lucide.dev/)
</div>

---

## 📋 Overview

The RVCE Counsellor Data Combiner is a sophisticated web-based tool designed specifically for R.V. College of Engineering to streamline the process of combining student–counsellor information from multiple Excel files, across different sheets within those files, and across different academic years. This tool maintains data integrity while providing professional formatting and comprehensive processing capabilities.

## Key Features

### 🔄 **Multi-Year Data Processing**
- **Year 4 (2022-2026)** - Final year students
- **Year 3 (2023-2027)** - Third year students  
- **Year 2 (2024-2028)** - Second year students

### 📊 **Advanced Excel Handling**
- **Hyperlink Extraction** - Automatically extracts email addresses from Excel hyperlinks
- **Rich Text Processing** - Handles complex Excel formatting and rich text objects
- **Formula Recognition** - Processes Excel formulas and their results
- **Branch Normalization** - Standardizes department names across all files
- **Robust Header Aliasing** - Flexible mapping for misspellings, spacing, and punctuation in Excel headers
- **Hyperlink Processing** - Handles both mailto: and tel: links
- **Bug Fixes** - Improved error handling and column alignment warnings resolved

### **Professional Output Formatting**
- **RVCE Logo Integration** - Embedded institutional branding
- **Color-coded Sections** - Visual hierarchy with batch and branch separators
- **Responsive Layout** - Optimized column widths and professional styling
- **Data Validation** - Comprehensive error checking and data cleaning

### 📱 **Cross-Platform Compatibility**
- **Responsive Design** - Works seamlessly on desktop, tablet, and mobile devices
- **Modern UI/UX** - Clean, intuitive interface built with React and Tailwind CSS
- **Real-time Processing** - Live feedback during file processing

## 🚀 Getting Started

### Prerequisites

- Node.js (v16 or higher)
- npm or yarn package manager
- Modern web browser

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/your-username/counsellors-rvce.git
   cd counsellors-rvce
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Start the development server**
   ```bash
   npm run dev
   ```

4. **Open your browser**
   ```
   http://localhost:3000
   ```

### Building for Production

```bash
npm run build
npm run preview
```

## 💻 Usage Guide

### Step 1: Upload Excel Files
- Click on the respective year cards (Year 2, 3, or 4)
- Select the corresponding `.xlsx` files
- Files are automatically validated upon upload

### Step 2: Process Data
- Click **"Combine Data"** to start processing
- Monitor real-time progress with the loading indicator
- View processing statistics upon completion

### Step 3: Download Results
- Click **"Download Excel"** to get the formatted output
- The file includes professional formatting with RVCE branding
- Data is organized by batch and branch with proper separators

## 🏗️ Technical Architecture

### Frontend Stack
- **React 18** - Modern component-based architecture
- **Vite** - Fast build tool and development server
- **Tailwind CSS** - Utility-first CSS framework
- **Lucide React** - Beautiful, customizable icons

### Core Libraries
- **ExcelJS** - Advanced Excel file processing
- **Google Fonts** - Professional typography (Inter font family)

### File Structure
```
counsellors-rvce/
├── public/
│   ├── index.html
│   └── assets/
├── src/
│   ├── App.jsx          # Main application component
│   ├── main.jsx         # Application entry point
│   ├── index.css        # Global styles
│   └── assets/
├── package.json
├── vite.config.js
├── tailwind.config.js
└── README.md
```

## 📊 Data Processing Pipeline

### 1. **File Validation**
- Format verification (`.xlsx` only)
- Sheet structure validation
- Header row detection
- Data integrity checks

### 2. **Data Extraction**
- **Hyperlink Processing** - Extracts emails from Excel hyperlinks
- **Cell Value Extraction** - Handles complex Excel objects
- **Branch Normalization** - Standardizes department names
- **Batch Detection** - Automatically identifies academic years

### 3. **Data Cleaning**
- Removes empty rows and invalid entries
- Standardizes data formats
- Validates USN patterns
- Cleans phone numbers and email addresses

### 4. **Output Generation**
- Professional Excel formatting
- RVCE logo integration
- Color-coded batch and branch sections
- Optimized column widths and styling

## 🎨 Supported Branch Mappings

| Input Format | Standardized Output |
|--------------|-------------------|
| CSE(AIML) | Computer Science Engineering (AI & ML) |
| AIML | Artificial Intelligence & Machine Learning |
| CSE | Computer Science Engineering |
| Data Science | CSE (Data Science) |
| Cyber Security | CSE (Cyber Security) |
| Aerospace Eng. | Aerospace Engineering |
| Civil Eng. | Civil Engineering |
| Chemical Eng. | Chemical Engineering |
| Mechanical Eng. | Mechanical Engineering |
| Information Science | Information Science & Engineering |
| EEE | Electrical & Electronics Engineering |
| ECE | Electronics & Communication Engineering |
| EIE | Electronics & Instrumentation Engineering |
| ET | Electronics & Telecommunication Engineering |
| IEM | Industrial Engineering & Management |

## 📈 Performance Optimizations

- **Code Splitting** - Vendor libraries separated for faster loading
- **Asset Optimization** - Compressed images and minified CSS/JS
- **Lazy Loading** - Components loaded on demand
- **Memory Management** - Efficient Excel processing with garbage collection

## 🔧 Configuration

### Environment Variables
Create a `.env` file in the root directory:
```env
VITE_APP_TITLE="RVCE Counsellor Data Combiner"
VITE_APP_VERSION="1.0.0"
```

## 🤝 Contributing

We welcome contributions from the RVCE community! Please follow these guidelines:

1. **Fork the repository**
2. **Create a feature branch** (`git checkout -b feature/amazing-feature`)
3. **Commit your changes** (`git commit -m 'Add amazing feature'`)
4. **Push to the branch** (`git push origin feature/amazing-feature`)
5. **Open a Pull Request**

### Development Guidelines
- Maintain responsive design principles
- Follow the existing code style

## 📄 License

This project is developed for R.V. College of Engineering and is intended for internal institutional use.

## 👥 Support & Contact
- :octocat: [Vivaan Hooda](https://github.com/VivaanHooda) | 📧 [Email](mailto:vivaan.hooda@gmail.com)
