# TSETMC to Excel Data Extractor

A robust C# application designed to fetch real-time stock market data from the Iranian stock exchange (TSETMC) and seamlessly export it directly into Microsoft Excel spreadsheets.

## 📌 Overview

Monitoring the Iranian stock market requires fast and organized access to data. This tool automates the process of extracting market data—ranging from general market overviews to specific, individual stock details—from the TSETMC platform. It dynamically writes the fetched data into Excel, making it easier for traders, analysts, and developers to track and analyze market trends.

## ✨ Key Features

- **Real-Time Data Fetching:** Extracts live market and specific stock data directly from TSETMC.
- **Excel Integration:** Automatically populates and updates data within an Excel file.
- **Adjustable Update Rate:** You can easily configure and change the refresh interval (update rate) to fetch data as fast or as slow as you need.
- **Customizable Fields:** Gives you full control over which specific data fields (columns) to include or exclude in your Excel output, ensuring you only see the data that matters to you.
- **Built with C#:** Developed using C# for high performance and reliable execution.

## 🚀 How It Works

1. The application connects to the TSETMC platform.
2. It fetches the requested data based on your selected fields.
3. The data is instantly pushed into the designated Excel file.
4. The cycle repeats based on your configured update rate.

## 🛠️ Tech Stack

- **Language:** C#
- **Target Platform:** Windows / Microsoft Excel
- **Data Source:** Tehran Securities Exchange Technology Management Co. (TSETMC)
