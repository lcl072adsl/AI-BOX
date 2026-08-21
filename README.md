<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>物流站所累計分析與配結率監控系統</title>
    <!-- Tailwind CSS -->
    <script src="https://tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cloudflare.com"/>
    <!-- Chart.js (Pinned Stable UMD Version) -->
    <script src="https://jsdelivr.net"></script>
    <!-- Chart.js DataLabels Plugin (Pinned Version) -->
    <script src="https://jsdelivr.net"></script>
    <!-- SheetJS for Excel / CSV -->
    <script src="https://jsdelivr.net"></script>
    <!-- PDF.js for PDF Rendering -->
    <script src="https://cloudflare.com"></script>
    
    <!-- 🟢 核心引入：手機與電腦全面相容的 Firebase 雲端連線外掛 -->
    <script src="https://gstatic.com"></script>
    <script src="https://gstatic.com"></script>
    
    <script>
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cloudflare.com';
        
        if (typeof ChartDataLabels !== 'undefined' && typeof Chart !== 'undefined') {
            Chart.unregister(ChartDataLabels);
        }

        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        brand: {
                            50: '#f0f7ff',
                            100: '#e0effe',
                            500: '#2563eb',
                            600: '#1d4ed8',
                            700: '#1e40af',
                            800: '#1e3a8a',
                        }
                    }
                }
            }
        }
    </script>
    <style>
        @import url('https://googleapis.com');
        body {
            font-family: 'Inter', system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background-color: #f8fafc;
        }
        .glass-card {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(8px);
            border: 1px solid rgba(226, 232, 240, 0.8);
        }
        .animate-fade-in {
            animation: fadeIn 0.25s ease-in-out;
        }
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(4px); }
            to { opacity: 1; transform: translateY(0); }
        }
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f5f9;
        }
        ::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
        }
    </style>
</head>
<body class="text-slate-800 antialiased min-h-screen flex flex-col">

    <header class="bg-slate-900 text-white sticky top-0 z-30 shadow-md">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
            <div class="flex items-center space-x-3">
                <i class="fa-solid fa-truck-fast text-blue-400 text-xl"></i>
                <h1 class="text-base sm:text-lg font-bold tracking-tight">物流站所累計分析與配結率監控系統</h1>
                <!-- 🟢 雲端同步動態指示燈 -->
                <span id="cloudStatusBadge" class="hidden sm:inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-semibold bg-indigo-500/20 text-indigo-300 border border-indigo-500/30">
                    <i id="cloudStatusIcon" class="fa-solid fa-cloud-arrow-down animate-pulse text-indigo-400"></i>
                    <span id="cloudStatusText">正在連接雲端...</span>
                </span>
            </div>
            <div class="flex items-center space-x-2 sm:space-x-3">
                <button onclick="openModal('dbManageModal')" class="bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm px-3.5 py-2 rounded-lg font-medium transition flex items-center gap-2 shadow-sm">
                    <i class="fa-solid fa-database"></i>
                    <span>資料庫管理 / 匯入</span>
                </button>
                <button onclick="openModal('uploadModal')" class="bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm px-3.5 py-2 rounded-lg font-medium transition flex items-center gap-2 shadow-sm">
                    <i class="fa-solid fa-cloud-arrow-up"></i>
                    <span>批量上傳 (多檔/AI辨識)</span>
                </button>
                <button onclick="exportToCSV()" class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm px-3.5 py-2 rounded-lg font-medium transition flex items-center gap-2 shadow-sm">
                    <i class="fa-solid fa-file-excel"></i>
                    <span>匯出 CSV</span>
                </button>
            </div>
        </div>
    </header>

    <!-- Toast Notification Banner -->
    <div id="toastNotification" class="fixed top-20 right-5 z-50 hidden bg-slate-800 text-white px-4 py-3 rounded-xl shadow-xl flex items-center gap-3 border border-slate-700 animate-fade-in max-w-md">
        <i id="toastIcon" class="fa-solid fa-circle-check text-emerald-400 text-lg"></i>
        <span id="toastMessage" class="text-xs font-medium">系統通知</span>
    </div>

    <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">

        <!-- Controls & Date Selector Bar -->
        <div class="glass-card rounded-xl p-4 shadow-sm flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
            <div class="flex flex-wrap items-center gap-4">
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">分析日期 / 累計區間</label>
                    <select id="dateSelect" onchange="filterData()" class="bg-white border border-slate-300 text-slate-800 text-sm rounded-lg focus:ring-blue-500 focus:border-blue-500 block px-3 py-1.5 shadow-sm font-medium">
                        <option value="ALL">全部收錄日期 (歷史累計分析)</option>
                    </select>
                </div>

                <!-- Section Department Filters -->
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">課別快速切換</label>
                    <div class="flex flex-wrap gap-1" id="sectionTagContainer">
                        <button onclick="setSectionFilter('ALL')" class="section-tag-btn active px-2.5 py-1 rounded-md text-xs font-semibold bg-blue-600 text-white shadow-xs">全部課別</button>
                        <button onclick="setSectionFilter('北一課')" class="section-tag-btn px-2.5 py-1 rounded-md text-xs font-medium bg-slate-100 text-slate-700 hover:bg-slate-200">北一課</button>
                        <button onclick="setSectionFilter('北二課')" class="section-tag-btn px-2.5 py-1 rounded-md text-xs font-medium bg-slate-100 text-slate-700 hover:bg-slate-200">北二課</button>
                        <button onclick="setSectionFilter('北三課')" class="section-tag-btn px-2.5 py-1 rounded-md text-xs font-medium bg-slate-100 text-slate-700 hover:bg-slate-200">北三課</button>
                        <button onclick="setSectionFilter('中課')" class="section-tag-btn px-2.5 py-1 rounded-md text-xs font-medium bg-slate-100 text-slate-700 hover:bg-slate-200">中課</button>
                        <button onclick="setSectionFilter('南一課')" class="section-tag-btn px-2.5 py-1 rounded-md text-xs font-medium bg-slate-100 text-slate-700 hover:bg-slate-200">南一課</button>
                        <button onclick="setSectionFilter('南二課')" class="section-tag-btn px-2.5 py-1 rounded-md text-xs font-medium bg-slate-100 text-slate-700 hover:bg-slate-200">南二課</button>
                        <button onclick="setSectionFilter('DC專案')" class="section-tag-btn px-2.5 py-1 rounded-md text-xs font-medium bg-slate-100 text-slate-700 hover:bg-slate-200">DC專案</button>
                    </div>
                </div>

                <!-- Type Filter (自配 / 委外) -->
                <div>
                    <label class="block text-xs font-semibold text-slate-500 mb-1">屬性別</label>
                    <div class="flex gap-1" id="typeTagContainer">
                        <button onclick="setTypeFilter('ALL')" class="type-tag-btn active px-2.5 py-1 rounded-md text-xs font-semibold bg-indigo-600 text-white shadow-xs">全部</button>
                        <button onclick="setTypeFilter('自配')" class="type-tag-btn px-2.5 py-1 rounded-md text-xs font-medium bg-slate-100 text-slate-700 hover:bg-slate-200">自配</button>
                        <button onclick="setTypeFilter('委外')" class="type-tag-btn px-2.5 py-1 rounded-md text-xs font-medium bg-slate-100 text-slate-700 hover:bg-slate-200">委外</button>
                    </div>
                </div>
            </div>

            <!-- Toggle Benchmark Panel Button -->
            <div class="flex items-center gap-2">
                <button id="toggleBenchBtn" onclick="toggleBenchmarkPanel()" class="px-3.5 py-1.5 rounded-lg text-xs font-semibold transition border flex items-center gap-1.5 bg-slate-800 text-white border-slate-700 shadow-sm">
                    <i class="fa-solid fa-sliders"></i>
                    <span>控制面板選項</span>
                </button>
            </div>
        </div>

        <!-- 數據展示主儀表板 -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div class="glass-card rounded-xl p-6 shadow-sm min-h-[350px]">
                <h3 class="text-xs font-bold text-slate-500 mb-4"><i class="fa-solid fa-chart-pie mr-2 text-blue-500"></i>配結率監控矩陣</h3>
                <div class="relative w-full h-64">
