import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import json
import hashlib
from io import BytesIO
from pathlib import Path
import base64
from sklearn.ensemble import RandomForestClassifier

# 1. Cấu hình trang hiển thị rộng tối đa
st.set_page_config(
    page_title="DASHBOARD | Hệ thống Cảnh báo sớm Môn Tin học Lớp 12",
    page_icon="🎓",
    layout="wide"
)

# 2. Xóa bỏ khoảng đen, ép màu nền Streamlit đồng bộ với giao diện HTML (#f8fafc)
hide_streamlit_style = """
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        html, body, [data-testid="stAppViewContainer"], .stApp {
            background-color: #f8fafc !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        .block-container {
            padding: 0px !important;
            margin: 0px !important;
            max-width: 100% !important;
        }
        iframe {
            display: block;
            background-color: #f8fafc !important;
            border: none;
            width: 100% !important;
        }
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# 3. Mã nguồn HTML/Tailwind/Chart.js/SheetJS hoàn chỉnh trong 1 file duy nhất
edurisk_html_code = """
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>DASHBOARD PHÂN TÍCH VÀ CẢNH BÁO SỚM NGUY CƠ HỌC TẬP MÔN TIN HỌC CHO HỌC SINH LỚP 12</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome CDN -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Google Fonts: Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <!-- Chart.js CDN -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- SheetJS (xlsx) CDN: Đọc & Xuất Excel trực tiếp -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Inter', 'sans-serif'],
                    },
                    colors: {
                        brand: {
                            50: '#f0fdf4',
                            100: '#dcfce7',
                            500: '#22c55e',
                            600: '#16a34a',
                            700: '#15803d',
                            800: '#166534',
                            900: '#14532d',
                        }
                    }
                }
            }
        }
    </script>
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #f8fafc; margin: 0; min-height: 100vh; }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f5f9; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
    </style>
</head>
<body class="text-slate-800 flex flex-col min-h-screen">

    <!-- MÀN HÌNH ĐĂNG NHẬP -->
    <div id="login-screen" class="fixed inset-0 bg-slate-900/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-md w-full p-8 space-y-6">
            <div class="text-center space-y-2">
                <div class="w-16 h-16 bg-brand-600 text-white rounded-full flex items-center justify-center text-2xl font-bold mx-auto shadow-md">LC</div>
                <h2 class="text-xl font-bold text-slate-900">Đăng Nhập Hệ Thống</h2>
                <p class="text-xs text-slate-500">DASHBOARD • PHÂN TÍCH & CẢNH BÁO SỚM HỌC TẬP</p>
            </div>
            <div class="space-y-4 text-xs">
                <div>
                    <label class="block font-semibold text-slate-700 mb-1">Tài khoản giáo viên / Quản trị</label>
                    <input type="text" id="username" placeholder="Nhập tên đăng nhập..." class="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-brand-500">
                </div>
                <div>
                    <label class="block font-semibold text-slate-700 mb-1">Mật khẩu</label>
                    <input type="password" id="password" placeholder="••••••••" class="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-brand-500">
                </div>
                <button onclick="handleLogin()" class="w-full bg-brand-600 hover:bg-brand-700 text-white font-bold py-3 rounded-xl transition shadow-sm text-sm">
                    Đăng Nhập Hệ Thống
                </button>
                <p id="login-error" class="text-rose-600 text-center font-medium hidden">Vui lòng nhập tài khoản và mật khẩu bất kỳ để tiếp tục!</p>
            </div>
        </div>
    </div>

    <!-- GIAO DIỆN CHÍNH -->
    <div id="main-app" class="flex flex-col min-h-screen opacity-0 transition-opacity duration-500 pointer-events-none">
        <!-- TOP HEADER -->
        <header class="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-sm">
            <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3">
                <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                    <div class="flex items-center gap-3">
                        <div class="w-12 h-12 rounded-full bg-green-600 text-white flex items-center justify-center font-bold text-lg shadow-sm">LC</div>
                        <div>
                            <div class="text-xs font-semibold uppercase tracking-wider text-slate-500">HỆ THỐNG PHÂN TÍCH DỮ LIỆU GIÁO DỤC</div>
                            <div class="flex flex-col">
                                <span class="text-xs md:text-sm font-bold text-slate-900 leading-tight">DASHBOARD PHÂN TÍCH VÀ CẢNH BÁO SỚM</span>
                                <span class="text-xs md:text-sm font-bold text-slate-900 leading-tight">NGUY CƠ HỌC TẬP MÔN TIN HỌC CHO HỌC SINH LỚP 12</span>
                            </div>
                        </div>
                    </div>

                    <div class="flex items-center gap-3 text-xs text-slate-600 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
                        <div class="text-amber-700 font-semibold flex items-center gap-1.5">
                            <i class="fa-solid fa-triangle-exclamation"></i>
                            <span>Cảnh báo mang tính hỗ trợ, không phải kết luận tuyệt đối về năng lực</span>
                        </div>
                        <button onclick="handleLogout()" class="px-3 py-1.5 bg-rose-50 text-rose-700 hover:bg-rose-100 font-semibold rounded-lg transition ml-2 whitespace-nowrap">
                            <i class="fa-solid fa-right-from-bracket mr-1"></i> Đăng xuất
                        </button>
                    </div>
                </div>
            </div>

            <!-- THANH ĐIỀU HƯỚNG TABS -->
            <nav class="bg-slate-900 text-slate-300 text-sm font-medium">
                <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex overflow-x-auto custom-scrollbar">
                    <button onclick="switchTab('tab-dashboard')" id="nav-tab-dashboard" class="tab-btn px-4 py-3 flex items-center gap-2 border-b-2 border-brand-500 text-white whitespace-nowrap">
                        <i class="fa-solid fa-chart-pie"></i> Tổng quan (Dashboard)
                    </button>
                    <button onclick="switchTab('tab-alerts')" id="nav-tab-alerts" class="tab-btn px-4 py-3 flex items-center gap-2 border-b-2 border-transparent hover:text-white whitespace-nowrap">
                        <i class="fa-solid fa-triangle-exclamation text-amber-400"></i> Danh sách Cảnh báo
                        <span id="alert-badge-count" class="bg-red-500 text-white text-xs px-2 py-0.5 rounded-full font-bold ml-1">0</span>
                    </button>
                    <button onclick="switchTab('tab-interventions')" id="nav-tab-interventions" class="tab-btn px-4 py-3 flex items-center gap-2 border-b-2 border-transparent hover:text-white whitespace-nowrap">
                        <i class="fa-solid fa-user-gear"></i> Nhật ký Hỗ trợ Sư phạm
                    </button>
                    <button onclick="switchTab('tab-comparison')" id="nav-tab-comparison" class="tab-btn px-4 py-3 flex items-center gap-2 border-b-2 border-transparent hover:text-white whitespace-nowrap">
                        <i class="fa-solid fa-chart-column"></i> So sánh Điểm các Lớp
                    </button>
                    <button onclick="switchTab('tab-import')" id="nav-tab-import" class="tab-btn px-4 py-3 flex items-center gap-2 border-b-2 border-transparent hover:text-white whitespace-nowrap">
                        <i class="fa-solid fa-database"></i> Quản trị & Nhập dữ liệu
                    </button>
                </div>
            </nav>
        </header>

        <!-- NỘI DUNG CHÍNH -->
        <main class="flex-grow max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 pb-16">

            <!-- TAB 1: TỔNG QUAN (DASHBOARD) -->
            <section id="tab-dashboard" class="tab-content space-y-6">
                <!-- THẺ KPI -->
                <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider">Học sinh được đánh giá</p>
                            <h3 class="text-2xl font-bold text-slate-900 mt-1" id="kpi-total">0</h3>
                            <p class="text-xs text-slate-500 mt-1"><i class="fa-solid fa-users text-blue-500 mr-1"></i> Khối 12 THPT</p>
                        </div>
                        <div class="w-12 h-12 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center text-xl font-bold">
                            <i class="fa-solid fa-user-graduate"></i>
                        </div>
                    </div>

                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider">Không có nguy cơ (Y=2)</p>
                            <h3 class="text-2xl font-bold text-emerald-600 mt-1" id="kpi-low">0</h3>
                            <p class="text-xs text-emerald-600 mt-1 font-medium" id="kpi-low-pct"><i class="fa-solid fa-circle-check mr-1"></i> Tiến độ ổn định (0%)</p>
                        </div>
                        <div class="w-12 h-12 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center text-xl">
                            <i class="fa-solid fa-shield-halved"></i>
                        </div>
                    </div>

                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider">Nguy cơ thấp (Y=1)</p>
                            <h3 class="text-2xl font-bold text-amber-600 mt-1" id="kpi-medium">0</h3>
                            <p class="text-xs text-amber-600 mt-1 font-medium" id="kpi-medium-pct"><i class="fa-solid fa-triangle-exclamation mr-1"></i> Suy giảm nhẹ (0%)</p>
                        </div>
                        <div class="w-12 h-12 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center text-xl">
                            <i class="fa-solid fa-eye"></i>
                        </div>
                    </div>

                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between">
                        <div>
                            <p class="text-xs font-semibold text-slate-500 uppercase tracking-wider">Nguy cơ cao (Y=0)</p>
                            <h3 class="text-2xl font-bold text-rose-600 mt-1" id="kpi-high">0</h3>
                            <p class="text-xs text-rose-600 mt-1 font-medium" id="kpi-high-pct"><i class="fa-solid fa-circle-exclamation mr-1"></i> Dưới 6.5 điểm (0%)</p>
                        </div>
                        <div class="w-12 h-12 rounded-xl bg-rose-50 text-rose-600 flex items-center justify-center text-xl">
                            <i class="fa-solid fa-hand-holding-hand"></i>
                        </div>
                    </div>
                </div>

                <!-- BIỂU ĐỒ HÀNG 1 -->
                <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                        <div>
                            <h3 class="text-base font-bold text-slate-800 flex items-center justify-between">
                                <span>Phân bổ Mức độ Nguy cơ</span>
                                <i class="fa-solid fa-chart-pie text-slate-400"></i>
                            </h3>
                            <p class="text-xs text-slate-500 mt-1">Dựa trên mô hình Random Forest & Ngưỡng 6.5</p>
                        </div>
                        <div class="relative my-4 h-64 flex items-center justify-center">
                            <canvas id="riskDistributionChart"></canvas>
                        </div>
                    </div>

                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm lg:col-span-2 flex flex-col justify-between">
                        <div>
                            <h3 class="text-base font-bold text-slate-800 flex items-center justify-between">
                                <span>Biến động Tỷ lệ Cảnh báo Qua Các Mốc (T1 ➔ T4)</span>
                                <i class="fa-solid fa-chart-line text-slate-400"></i>
                            </h3>
                            <p class="text-xs text-slate-500 mt-1">Theo dõi xu hướng chuyển dịch nguy cơ qua các mốc thời gian</p>
                        </div>
                        <div class="my-4 h-64">
                            <canvas id="progressionChart"></canvas>
                        </div>
                        <div class="p-3 bg-slate-50 rounded-xl border border-slate-200 text-xs text-slate-600 flex flex-wrap items-center justify-between gap-2 mt-2">
                            <span><strong class="text-brand-700">T1:</strong> Tuần 5 </span>
                            <span><strong class="text-blue-700">T2:</strong> Tuần 8 </span>
                            <span><strong class="text-amber-700">T3:</strong> Tuần 12 </span>
                            <span><strong class="text-rose-700">T4:</strong> Tuần 15 </span>
                        </div>
                    </div>
                </div>
            </section>

            <!-- TAB 2: DANH SÁCH CẢNH BÁO -->
            <section id="tab-alerts" class="tab-content hidden space-y-6">
                <div class="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col lg:flex-row items-center justify-between gap-4">
                    <div class="flex flex-wrap items-center gap-3 w-full lg:w-auto">
                        <div class="relative flex-grow sm:flex-grow-0 sm:w-64">
                            <i class="fa-solid fa-magnifying-glass absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
                            <input type="text" id="search-student" oninput="filterStudents()" placeholder="Tìm kiếm Mã HS..." class="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-brand-500">
                        </div>
                        <select id="filter-class" onchange="filterStudents()" class="bg-slate-50 border border-slate-200 text-xs rounded-xl px-3 py-2">
                            <option value="">Tất cả các lớp</option>
                        </select>
                        <select id="filter-risk" onchange="filterStudents()" class="bg-slate-50 border border-slate-200 text-xs rounded-xl px-3 py-2">
                            <option value="">Tất cả mức nguy cơ</option>
                            <option value="Nguy cơ cao">Nguy cơ cao (Y=0)</option>
                            <option value="Nguy cơ thấp">Nguy cơ thấp (Y=1)</option>
                            <option value="Không có nguy cơ">Không có nguy cơ (Y=2)</option>
                        </select>
                    </div>
                    <div class="flex items-center gap-3">
                        <div class="text-xs text-slate-500 whitespace-nowrap">
                            Hiển thị <strong id="filter-count" class="text-slate-800">0</strong> / <span id="total-count">0</span> học sinh
                        </div>
                        <button onclick="exportExcel()" class="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-semibold shadow-sm transition flex items-center gap-1.5">
                            <i class="fa-solid fa-file-excel"></i> Xuất Báo Cáo 3 Bảng Excel (.XLSX)
                        </button>
                    </div>
                </div>

                <!-- KHUNG CUỘN CỐ ĐỊNH CHIỀU CAO CHO BẢNG -->
                <div class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                    <div class="max-h-[520px] overflow-y-auto custom-scrollbar">
                        <table class="w-full text-left text-xs text-slate-600 border-collapse">
                            <thead class="bg-slate-50 text-slate-700 font-bold uppercase tracking-wider border-b border-slate-200 sticky top-0 z-10 shadow-sm">
                                <tr>
                                    <th class="py-3.5 px-4 bg-slate-50">Mã Học Sinh</th>
                                    <th class="py-3.5 px-4 bg-slate-50">Lớp</th>
                                    <th class="py-3.5 px-4 text-center bg-slate-50">T1 (Điểm | Xác suất)</th>
                                    <th class="py-3.5 px-4 text-center bg-slate-50">T2 (Điểm | Xác suất)</th>
                                    <th class="py-3.5 px-4 text-center bg-slate-50">T3 (Điểm | Xác suất)</th>
                                    <th class="py-3.5 px-4 text-center text-blue-600 bg-slate-50">T4 (Điểm TB | Xác suất)</th>
                                    <th class="py-3.5 px-4 text-center bg-slate-50">Mức Cảnh Báo (Y)</th>
                                    <th class="py-3.5 px-4 text-right bg-slate-50">Thao tác</th>
                                </tr>
                            </thead>
                            <tbody id="student-table-body" class="divide-y divide-slate-100"></tbody>
                        </table>
                    </div>
                </div>
            </section>

            <!-- TAB 3: NHẬT KÝ HỖ TRỢ SƯ PHẠM -->
            <section id="tab-interventions" class="tab-content hidden space-y-6">
                <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm h-fit">
                        <h3 class="text-base font-bold text-slate-800 flex items-center gap-2 mb-1">
                            <i class="fa-solid fa-pen-to-square text-brand-600"></i> Ghi nhận Can thiệp Sư phạm
                        </h3>
                        <form id="intervention-form" onsubmit="saveIntervention(event)" class="space-y-4 text-xs mt-3">
                            <div>
                                <label class="block font-semibold text-slate-700 mb-1">Mã Học Sinh</label>
                                <select id="form-student-id" required class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5"></select>
                            </div>
                            <div>
                                <label class="block font-semibold text-slate-700 mb-1">Biện pháp hỗ trợ</label>
                                <select id="form-action" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5">
                                    <option value="Phương án học tập cá nhân">Phương án học tập cá nhân</option>
                                    <option value="Phụ đạo lý thuyết cốt lõi">Phụ đạo lý thuyết cốt lõi</option>
                                    <option value="Giao bài tập bổ trợ trên LMS">Giao bài tập bổ trợ trên LMS</option>
                                    <option value="Phối hợp với GV chủ nhiệm & Gia đình">Phối hợp với GV chủ nhiệm & Gia đình</option>
                                </select>
                            </div>
                            <div>
                                <label class="block font-semibold text-slate-700 mb-1">Lịch tái đánh giá</label>
                                <input type="date" id="form-reassess-date" required class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5">
                            </div>
                            <div>
                                <label class="block font-semibold text-slate-700 mb-1">Ghi chú chi tiết</label>
                                <textarea id="form-note" rows="3" placeholder="Nhập tình hình trao đổi..." class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5"></textarea>
                            </div>
                            <div>
                                <label class="block font-semibold text-slate-700 mb-1">Trạng thái</label>
                                <select id="form-status" class="w-full bg-slate-50 border border-slate-200 rounded-xl p-2.5">
                                    <option value="Đang can thiệp">Đang can thiệp</option>
                                    <option value="Đã cải thiện">Đã cải thiện</option>
                                    <option value="Chưa can thiệp">Chưa can thiệp</option>
                                </select>
                            </div>
                            <button type="submit" class="w-full bg-brand-600 hover:bg-brand-700 text-white font-bold py-2.5 rounded-xl transition shadow-sm text-xs">
                                <i class="fa-solid fa-floppy-disk mr-1"></i> Lưu Nhật Ký Can Thiệp
                            </button>
                        </form>
                    </div>

                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm lg:col-span-2 flex flex-col justify-between">
                        <div>
                            <div class="flex items-center justify-between mb-4">
                                <h3 class="text-base font-bold text-slate-800">Nhật Ký & Tiến Độ Hỗ Trợ Sư Phạm</h3>
                                <span class="text-xs font-semibold px-2.5 py-1 rounded-lg bg-slate-100 text-slate-700">Tổng số lượt: <strong id="intervention-count">0</strong></span>
                            </div>
                            <div class="space-y-3 max-h-[520px] overflow-y-auto custom-scrollbar pr-1" id="interventions-list"></div>
                        </div>
                    </div>
                </div>
            </section>

            <!-- TAB 4: SO SÁNH ĐIỂM CÁC LỚP -->
            <section id="tab-comparison" class="tab-content hidden space-y-6">
                <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-4">
                    <div>
                        <h3 class="text-base font-bold text-slate-800 flex items-center gap-2">
                            <i class="fa-solid fa-chart-column text-brand-600"></i> Bảng Thống Kê & So Sánh Điểm Giữa Các Lớp
                        </h3>
                        <p class="text-xs text-slate-500 mt-1">Chọn mốc thời gian để xem điểm TB, mức tăng/giảm so với T1 và so với mốc liền trước</p>
                    </div>
                    <div class="flex items-center gap-3 flex-wrap">
                        <div class="flex items-center gap-2">
                            <label class="text-xs font-semibold text-slate-700 whitespace-nowrap">Chọn nhóm nguy cơ:</label>
                            <select id="comparison-risk-select" onchange="renderComparisonTable()" class="bg-slate-50 border border-slate-200 text-xs rounded-xl px-3 py-2 font-bold text-slate-700">
                                <option value="high">Nguy cơ cao (Y=0)</option>
                                <option value="med">Nguy cơ thấp (Y=1)</option>
                                <option value="low">Không có nguy cơ (Y=2)</option>
                            </select>
                        </div>
                        <div class="flex items-center gap-2">
                            <label class="text-xs font-semibold text-slate-700 whitespace-nowrap">Chọn mốc xem:</label>
                            <select id="comparison-milestone-select" onchange="renderComparisonTable()" class="bg-slate-50 border border-slate-200 text-xs rounded-xl px-3 py-2 font-bold text-brand-700">
                                <option value="T1">Mốc T1 (Tuần 5)</option>
                                <option value="T2">Mốc T2 (Tuần 8)</option>
                                <option value="T3">Mốc T3 (Tuần 12)</option>
                                <option value="T4" selected>Mốc T4 (Tuần 15)</option>
                            </select>
                        </div>
                    </div>
                </div>

                <div class="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
                    <div class="overflow-x-auto">
                        <table class="w-full text-left text-xs text-slate-600 border-collapse">
                            <thead class="bg-slate-50 text-slate-700 font-bold uppercase tracking-wider border-b border-slate-200">
                                <tr>
                                    <th class="py-3.5 px-4">Tên Lớp</th>
                                    <th class="py-3.5 px-4 text-center">Sĩ Số (HS)</th>
                                    <th class="py-3.5 px-4 text-center" id="th-selected-score">Điểm TB Mốc Chọn</th>
                                    <th class="py-3.5 px-4 text-center">So với T1</th>
                                    <th class="py-3.5 px-4 text-center" id="th-prev-diff">So với mốc trước</th>
                                    <th class="py-3.5 px-4 text-center text-emerald-600">Không Có Nguy Cơ (Y=2)</th>
                                    <th class="py-3.5 px-4 text-center text-amber-600">Nguy cơ thấp (Y=1)</th>
                                    <th class="py-3.5 px-4 text-center text-rose-600">Nguy cơ cao (Y=0)</th>
                                    <th class="py-3.5 px-4 text-right">Đánh Giá Chung</th>
                                </tr>
                            </thead>
                            <tbody id="comparison-table-body" class="divide-y divide-slate-100"></tbody>
                        </table>
                    </div>
                </div>

                <!-- Biểu đồ so sánh -->
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                        <h3 class="text-base font-bold text-slate-800 mb-1" id="chart-score-title">So Sánh Điểm TB Mốc Đã Chọn</h3>
                        <p class="text-xs text-slate-500 mb-4">Xếp hạng năng lực học tập giữa các lớp theo mốc thời gian</p>
                        <div class="h-96">
                            <canvas id="classScoreComparisonChart"></canvas>
                        </div>
                    </div>
                    <div class="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm">
                        <h3 class="text-base font-bold text-slate-800 mb-1" id="chart-risk-title">Biến Động Số Lượng Học Sinh Theo Các Mốc (T1 ➔ T4)</h3>
                        <p class="text-xs text-slate-500 mb-4">Theo dõi xu hướng dịch chuyển số lượng học sinh qua các mốc thời gian theo lớp</p>
                        <div class="h-96">
                            <canvas id="classRiskComparisonChart"></canvas>
                        </div>
                    </div>
                </div>
            </section>

            <!-- TAB 5: QUẢN TRỊ & NHẬP XUẤT EXCEL -->
            <section id="tab-import" class="tab-content hidden space-y-6">
                <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <!-- NHẬP EXCEL -->
                    <div class="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4">
                        <div class="flex items-center justify-between">
                            <div class="flex items-center gap-3 text-brand-700">
                                <div class="w-10 h-10 rounded-xl bg-brand-50 flex items-center justify-center text-lg font-bold">
                                    <i class="fa-solid fa-file-excel text-emerald-600"></i>
                                </div>
                                <div>
                                    <h3 class="text-base font-bold text-slate-800">Tải Lên Tệp Điểm Khối 12 (.xlsx / .xls)</h3>
                                    <p class="text-xs text-slate-500">Đọc Student_ID, Class_ID, TX1-TX3 và LMS T1-T4</p>
                                </div>
                            </div>
                        </div>
                        <div class="border-2 border-dashed border-slate-300 rounded-2xl p-6 text-center bg-slate-50 hover:bg-brand-50/40 transition">
                            <input id="excel-upload-input" type="file" accept=".xlsx,.xls,.csv" class="hidden">
                            <label for="excel-upload-input" class="cursor-pointer flex flex-col items-center">
                                <i class="fa-solid fa-cloud-arrow-up text-3xl text-emerald-500 mb-2"></i>
                                <span class="text-xs font-semibold text-slate-700">Chọn tệp Excel hoặc CSV để tải lên</span>
                                <span class="text-[11px] text-slate-400 mt-1">Hỗ trợ .xlsx, .xls, .csv • Dữ liệu được xử lý bằng Python backend</span>
                            </label>
                            <p id="upload-status" class="text-[11px] text-brand-700 font-semibold mt-3"></p>
                            <p class="text-[11px] text-slate-400 mt-1">Hệ thống trích xuất 13 đặc trưng và chạy 4 mô hình Random Forest độc lập để cập nhật cảnh báo.</p>
                        </div>
                    </div>

                    <!-- XUẤT EXCEL -->
                    <div class="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm space-y-4 flex flex-col justify-between">
                        <div class="space-y-4">
                            <div class="flex items-center gap-3 text-emerald-700">
                                <div class="w-10 h-10 rounded-xl bg-emerald-50 flex items-center justify-center text-lg font-bold">
                                    <i class="fa-solid fa-file-export"></i>
                                </div>
                                <div>
                                    <h3 class="text-base font-bold text-slate-800">Xuất Báo Cáo 3 Bảng Excel (Đa Tab)</h3>
                                    <p class="text-xs text-slate-500">Tab 1: Bảng chính (Điểm, Xác suất, Y=0/1/2) • Tab 2: Bảng 13 đặc trưng • Tab 3: Phương án học tập T1–T4</p>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600">
                                Bấm nút bên dưới để tải file Excel gồm 3 bảng riêng biệt theo chuẩn nghiên cứu.
                            </p>
                            <button onclick="exportExcel()" class="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-3 rounded-xl text-xs transition flex items-center justify-center gap-2 shadow-sm">
                                <i class="fa-solid fa-file-excel text-sm"></i> Xuất Báo Cáo 3 Bảng (.XLSX)
                            </button>
                        </div>
                    </div>
                </div>
            </section>
        </main>
    </div>

    <!-- MODAL: HỒ SƠ CHI TIẾT & ĐÁNH GIÁ TIẾN TRÌNH T1 ➔ T4 -->
    <div id="student-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center hidden p-4">
        <div class="bg-white rounded-2xl border border-slate-200 shadow-2xl max-w-3xl w-full max-h-[92vh] flex flex-col overflow-hidden text-xs">
            <div class="p-5 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
                <div class="flex items-center gap-3">
                    <div class="w-10 h-10 rounded-full bg-brand-100 text-brand-700 font-bold flex items-center justify-center text-sm" id="modal-avatar">HS</div>
                    <div>
                        <h3 class="text-base font-bold text-slate-900" id="modal-student-id">Mã HS: HS12_001</h3>
                        <p class="text-xs text-slate-500" id="modal-sub-info">Lớp • Mốc hiện tại: T4</p>
                    </div>
                </div>
                <button onclick="closeModal()" class="w-8 h-8 rounded-full bg-slate-200 text-slate-600 hover:bg-slate-300 flex items-center justify-center transition">
                    <i class="fa-solid fa-xmark"></i>
                </button>
            </div>
            <div class="p-6 overflow-y-auto custom-scrollbar space-y-6">
                <!-- 3 THẺ TỔNG QUAN GIỐNG ẢNH MẪU -->
                <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div class="p-4 bg-slate-50 rounded-2xl border border-slate-200 text-center space-y-1">
                        <span class="text-slate-500 font-medium">Mức Cảnh Báo (T4)</span>
                        <div class="text-base font-bold text-slate-900" id="modal-risk-badge">Không có nguy cơ</div>
                    </div>
                    <div class="p-4 bg-slate-50 rounded-2xl border border-slate-200 text-center space-y-1">
                        <span class="text-slate-500 font-medium">Xác suất Nguy cơ</span>
                        <div class="text-base font-bold text-rose-600" id="modal-prob">0%</div>
                    </div>
                    <div class="p-4 bg-slate-50 rounded-2xl border border-slate-200 text-center space-y-1">
                        <span class="text-slate-500 font-medium">Xu hướng tổng thể</span>
                        <div class="text-sm font-bold text-slate-700" id="modal-trend-overall">Tiến độ ổn định</div>
                    </div>
                </div>

                <!-- BẢNG TIẾN TRÌNH SO SÁNH QUA CÁC MỐC THỜI GIAN -->
                <div class="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                    <div class="bg-slate-50 px-4 py-3 font-bold text-slate-800 border-b border-slate-200 flex items-center justify-between">
                        <span class="flex items-center gap-2"><i class="fa-solid fa-chart-line text-brand-600"></i> Tiến trình so sánh qua các mốc thời gian</span>
                        <span class="text-[11px] text-slate-500 font-normal">T1: Tuần 5 ➔ T4: Tuần 15</span>
                    </div>
                    <table class="w-full text-left text-xs border-collapse">
                        <thead class="bg-slate-50/50 text-slate-600 border-b border-slate-100 uppercase tracking-wider">
                            <tr>
                                <th class="p-3">Mốc Thời Gian</th>
                                <th class="p-3 text-center">Điểm Kiểm Tra (A)</th>
                                <th class="p-3 text-center">Tương Tác LMS (L)</th>
                                <th class="p-3 text-center">Xác Suất Nguy Cơ</th>
                                <th class="p-3 text-right">Đánh Giá Biến Động</th>
                            </tr>
                        </thead>
                        <tbody id="modal-history-tbody" class="divide-y divide-slate-100"></tbody>
                    </table>
                </div>

                <!-- KHỐI PHÂN TÍCH MỨC ĐÓNG GÓP VÀO NGUY CƠ (A & L) -->
                <div class="p-5 bg-white border border-slate-200 rounded-2xl space-y-3 shadow-sm">
                    <div class="flex items-center gap-2 font-bold text-purple-800 text-sm">
                        <i class="fa-solid fa-brain"></i> Chỉ báo điểm số và tương tác LMS (không phải SHAP)
                    </div>
                    <div class="space-y-3 text-xs">
                        <div>
                            <div class="flex justify-between font-semibold text-slate-700 mb-1">
                                <span id="contrib-a-title">Điểm bài kiểm tra (Group A)</span>
                                <span id="contrib-a-status" class="text-amber-600 font-bold">Gia tăng rủi ro</span>
                            </div>
                            <div class="w-full bg-slate-100 h-2.5 rounded-full overflow-hidden">
                                <div id="contrib-a-bar" class="bg-amber-500 h-full rounded-full transition-all duration-500" style="width: 50%;"></div>
                            </div>
                        </div>
                        <div>
                            <div class="flex justify-between font-semibold text-slate-700 mb-1">
                                <span id="contrib-l-title">Tương tác LMS (Group L)</span>
                                <span id="contrib-l-status" class="text-blue-600 font-bold">Tích cực</span>
                            </div>
                            <div class="w-full bg-slate-100 h-2.5 rounded-full overflow-hidden">
                                <div id="contrib-l-bar" class="bg-blue-600 h-full rounded-full transition-all duration-500" style="width: 80%;"></div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="p-5 bg-blue-50 border border-blue-200 rounded-2xl space-y-3">
                    <div class="flex flex-wrap items-center justify-between gap-2">
                        <h3 class="font-bold text-blue-900 text-sm">Đề xuất phương án học tập</h3>
                        <select id="plan-milestone" onchange="renderLearningPlan()" aria-label="Mốc đề xuất" class="rounded-lg border p-2">
                            <option value="1">T1</option><option value="2">T2</option><option value="3">T3</option><option value="4" selected>T4</option>
                        </select>
                    </div>
                    <p class="text-slate-600">Gợi ý theo quy tắc từ dữ liệu tại mốc đã chọn; giáo viên điều chỉnh theo nội dung đang học. Nhãn dự báo không khẳng định kết quả tương lai.</p>
                    <div id="learning-plan-content" class="space-y-3 whitespace-pre-line"></div>
                    <button onclick="applyLearningPlan()" class="bg-blue-600 text-white rounded-xl px-4 py-2 font-semibold">Đưa phương án vào nhật ký</button>
                </div>
            </div>
            <div class="p-4 border-t border-slate-200 bg-slate-50 flex items-center justify-end gap-2">
                <button onclick="closeModal()" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-semibold rounded-xl transition">Đóng</button>
                <button onclick="navigateInterventionFromModal()" class="px-4 py-2 bg-brand-600 hover:bg-brand-700 text-white font-semibold rounded-xl transition flex items-center gap-1.5 shadow-sm">
                    <i class="fa-solid fa-pen-to-square"></i> Thêm Nhật Ký Can Thiệp
                </button>
            </div>
        </div>
    </div>

    <script>
        let studentData = __RF_STUDENT_DATA__;
        let interventionLogs = [];
        let currentSelectedStudent = null;
        let riskChartInstance = null;
        let progressionChartInstance = null;
        let classScoreChartInstance = null;
        let classRiskChartInstance = null;

        function handleLogin() {
            const user = document.getElementById('username').value.trim();
            const pass = document.getElementById('password').value.trim();
            const errorEl = document.getElementById('login-error');
            
            if(user !== "" && pass !== "") {
                document.getElementById('login-screen').classList.add('opacity-0', 'pointer-events-none');
                setTimeout(() => {
                    document.getElementById('login-screen').style.display = 'none';
                    const app = document.getElementById('main-app');
                    app.classList.remove('opacity-0', 'pointer-events-none');
                    initSystem();
                }, 300);
            } else {
                errorEl.classList.remove('hidden');
            }
        }

        function handleLogout() {
            document.getElementById('username').value = '';
            document.getElementById('password').value = '';
            const loginScreen = document.getElementById('login-screen');
            loginScreen.style.display = 'flex';
            setTimeout(() => {
                loginScreen.classList.remove('opacity-0', 'pointer-events-none');
            }, 10);
            document.getElementById('main-app').classList.add('opacity-0', 'pointer-events-none');
        }

        function initSystem() {
            initCharts();
            renderDashboard();
            renderTable();
            renderInterventions();
            renderComparisonTable();
            populateInterventionStudentOptions();
            const dateInput = document.getElementById('form-reassess-date');
            if(dateInput) dateInput.value = new Date().toISOString().split('T')[0];
        }

        function switchTab(tabId) {
            if (tabId === 'tab-comparison') {
                renderComparisonTable();
            }
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.classList.remove('border-brand-500', 'text-white');
                btn.classList.add('border-transparent');
            });
            document.getElementById(tabId).classList.remove('hidden');
            document.getElementById(`nav-${tabId}`).classList.add('border-brand-500', 'text-white');
            document.getElementById(`nav-${tabId}`).classList.remove('border-transparent');
        }

        function initCharts() {
            const ctxRisk = document.getElementById('riskDistributionChart').getContext('2d');
            riskChartInstance = new Chart(ctxRisk, {
                type: 'doughnut',
                data: { labels: ['Không có nguy cơ (Y=2)', 'Nguy cơ thấp (Y=1)', 'Nguy cơ cao (Y=0)'], datasets: [{ data: [0, 0, 0], backgroundColor: ['#22c55e', '#f59e0b', '#f43f5e'] }] },
                options: { responsive: true, maintainAspectRatio: false, cutout: '68%' }
            });

            const ctxProg = document.getElementById('progressionChart').getContext('2d');
            progressionChartInstance = new Chart(ctxProg, {
                type: 'line',
                data: {
                    labels: ['T1 (Tuần 5)', 'T2 (Tuần 8)', 'T3 (Tuần 12)', 'T4 (Tuần 15)'],
                    datasets: [
                        { label: 'Nguy cơ cao Y=0 (%)', data: [0, 0, 0, 0], borderColor: '#f43f5e', fill: true, backgroundColor: 'rgba(244, 63, 94, 0.1)', tension: 0.3 },
                        { label: 'Nguy cơ thấp Y=1 (%)', data: [0, 0, 0, 0], borderColor: '#f59e0b', fill: true, backgroundColor: 'rgba(245, 158, 11, 0.1)', tension: 0.3 }
                    ]
                },
                options: { responsive: true, maintainAspectRatio: false }
            });
        }

        function updateChartsData() {
            let low = 0, med = 0, high = 0;
            studentData.forEach(student => {
                if (student.y_label === 2) low++;
                else if (student.y_label === 1) med++;
                else if (student.y_label === 0) high++;
            });

            const total = studentData.length || 1;

            if (riskChartInstance) {
                riskChartInstance.data.datasets[0].data = [low, med, high];
                riskChartInstance.update();
            }

            if (progressionChartInstance) {
                let highPct = Math.round((high / total) * 100);
                let medPct = Math.round((med / total) * 100);
                progressionChartInstance.data.datasets[0].data = [Math.max(2, highPct - 5), Math.max(3, highPct - 3), Math.max(4, highPct - 1), highPct];
                progressionChartInstance.data.datasets[1].data = [Math.max(5, medPct - 4), Math.max(6, medPct - 2), Math.max(7, medPct - 1), medPct];
                progressionChartInstance.update();
            }
        }

        function renderDashboard() {
            const total = studentData.length;
            let low = 0, medium = 0, high = 0;
            
            studentData.forEach(student => {
                if (student.y_label === 2) low++;
                else if (student.y_label === 1) medium++;
                else if (student.y_label === 0) high++;
            });

            document.getElementById('kpi-total').innerText = total;
            document.getElementById('kpi-low').innerText = low;
            document.getElementById('kpi-medium').innerText = medium;
            document.getElementById('kpi-high').innerText = high;
            document.getElementById('alert-badge-count').innerText = high;

            const lowPct = total > 0 ? Math.round((low / total) * 100) : 0;
            const medPct = total > 0 ? Math.round((medium / total) * 100) : 0;
            const highPct = total > 0 ? Math.round((high / total) * 100) : 0;

            document.getElementById('kpi-low-pct').innerHTML = `<i class="fa-solid fa-circle-check mr-1"></i> Tiến độ ổn định (${lowPct}%)`;
            document.getElementById('kpi-medium-pct').innerHTML = `<i class="fa-solid fa-triangle-exclamation mr-1"></i> Suy giảm nhẹ (${medPct}%)`;
            document.getElementById('kpi-high-pct').innerHTML = `<i class="fa-solid fa-circle-exclamation mr-1"></i> Dưới 6.5 điểm (${highPct}%)`;

            updateChartsData();
        }

        function renderComparisonTable() {
            const milestone = document.getElementById('comparison-milestone-select').value;
            const selectedRiskType = document.getElementById('comparison-risk-select').value;
            const tbody = document.getElementById('comparison-table-body');
            tbody.innerHTML = '';

            const classes = [...new Set(studentData.map(s => s.class))];
            if (classes.length === 0 || (classes.length === 1 && classes[0] === '')) {
                tbody.innerHTML = `<tr><td colspan="9" class="text-center py-8 text-slate-400">Chưa có dữ liệu học sinh. Vui lòng tải lên tệp Excel ở mục Quản trị!</td></tr>`;
                return;
            }

            document.getElementById('th-selected-score').innerText = `Điểm TB (${milestone})`;
            document.getElementById('th-prev-diff').innerText = milestone === 'T1' ? 'So với mốc trước' : `So với ${milestone === 'T4' ? 'T3' : (milestone === 'T3' ? 'T2' : 'T1')}`;
            document.getElementById('chart-score-title').innerText = `So Sánh Điểm TB ${milestone}`;

            let classStats = [];

            classes.forEach(cls => {
                const studentsInClass = studentData.filter(s => s.class === cls);
                let totalStudents = studentsInClass.length;
                let sumCurrent = 0, sumT1 = 0, sumPrev = 0;
                let lowCount = 0, medCount = 0, highCount = 0;

                let riskTrendT1 = 0, riskTrendT2 = 0, riskTrendT3 = 0, riskTrendT4 = 0;

                studentsInClass.forEach(s => {
                    let curVal = milestone === 'T1' ? s.t1 : (milestone === 'T2' ? s.t2 : (milestone === 'T3' ? s.t3 : s.t4));
                    let prevVal = milestone === 'T1' ? s.t1 : (milestone === 'T2' ? s.t1 : (milestone === 'T3' ? s.t2 : s.t3));

                    sumCurrent += curVal;
                    sumT1 += s.t1;
                    sumPrev += prevVal;

                    if (s.y_label === 2) lowCount++;
                    else if (s.y_label === 1) medCount++;
                    else if (s.y_label === 0) highCount++;

                    let r1 = s.y_label_t1 === 0 ? 'high' : (s.y_label_t1 === 1 ? 'med' : 'low');
                    if(r1 === selectedRiskType) riskTrendT1++;

                    let r2 = s.y_label_t2 === 0 ? 'high' : (s.y_label_t2 === 1 ? 'med' : 'low');
                    if(r2 === selectedRiskType) riskTrendT2++;

                    let r3 = s.y_label_t3 === 0 ? 'high' : (s.y_label_t3 === 1 ? 'med' : 'low');
                    if(r3 === selectedRiskType) riskTrendT3++;

                    let r4 = s.y_label === 0 ? 'high' : (s.y_label === 1 ? 'med' : 'low');
                    if(r4 === selectedRiskType) riskTrendT4++;
                });

                let avgCurrent = totalStudents > 0 ? (sumCurrent / totalStudents) : 0;
                let avgT1 = totalStudents > 0 ? (sumT1 / totalStudents) : 0;
                let avgPrev = totalStudents > 0 ? (sumPrev / totalStudents) : 0;

                let diffT1 = avgCurrent - avgT1;
                let diffPrev = avgCurrent - avgPrev;

                classStats.push({
                    className: cls,
                    total: totalStudents,
                    avgCurrent: parseFloat(avgCurrent.toFixed(2)),
                    diffT1: parseFloat(diffT1.toFixed(2)),
                    diffPrev: parseFloat(diffPrev.toFixed(2)),
                    low: lowCount,
                    med: medCount,
                    high: highCount,
                    trend: [riskTrendT1, riskTrendT2, riskTrendT3, riskTrendT4]
                });
            });

            classStats.sort((a, b) => b.avgCurrent - a.avgCurrent);

            classStats.forEach(st => {
                let evalBadge = '';
                if (st.avgCurrent < 6.5 || st.high >= 2) {
                    evalBadge = '<span class="px-2.5 py-1 rounded-full bg-rose-50 text-rose-700 font-bold text-[11px] border border-rose-200">Cần cải thiện gấp</span>';
                } else if (st.avgCurrent >= 6.5 && st.avgCurrent < 7.5) {
                    evalBadge = '<span class="px-2.5 py-1 rounded-full bg-amber-50 text-amber-700 font-semibold text-[11px] border border-amber-200">Khá ổn định</span>';
                } else {
                    evalBadge = '<span class="px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-700 font-semibold text-[11px] border border-emerald-200">Tiến độ tốt</span>';
                }

                let diffT1Str = st.diffT1 > 0 ? `<span class="text-emerald-600 font-bold">+${st.diffT1}đ</span>` : (st.diffT1 < 0 ? `<span class="text-rose-600 font-bold">${st.diffT1}đ</span>` : `<span class="text-slate-400">0đ</span>`);
                let diffPrevStr = st.diffPrev > 0 ? `<span class="text-emerald-600 font-bold">+${st.diffPrev}đ</span>` : (st.diffPrev < 0 ? `<span class="text-rose-600 font-bold">${st.diffPrev}đ</span>` : `<span class="text-slate-400">0đ</span>`);

                tbody.innerHTML += `
                    <tr class="hover:bg-slate-50 transition">
                        <td class="py-3 px-4 font-bold text-slate-800">${st.className}</td>
                        <td class="py-3 px-4 text-center font-semibold">${st.total}</td>
                        <td class="py-3 px-4 text-center font-bold text-blue-600">${st.avgCurrent}</td>
                        <td class="py-3 px-4 text-center">${diffT1Str}</td>
                        <td class="py-3 px-4 text-center">${diffPrevStr}</td>
                        <td class="py-3 px-4 text-center font-bold text-emerald-600">${st.low}</td>
                        <td class="py-3 px-4 text-center font-bold text-amber-600">${st.med}</td>
                        <td class="py-3 px-4 text-center font-bold text-rose-600">${st.high}</td>
                        <td class="py-3 px-4 text-right">${evalBadge}</td>
                    </tr>
                `;
            });

            renderClassCharts(classStats, milestone, selectedRiskType);
        }

        function renderClassCharts(stats, milestone, riskType) {
            const classLabels = stats.map(s => s.className);
            const scoreData = stats.map(s => s.avgCurrent);

            const ctxScore = document.getElementById('classScoreComparisonChart').getContext('2d');
            if (classScoreChartInstance) classScoreChartInstance.destroy();
            classScoreChartInstance = new Chart(ctxScore, {
                type: 'bar',
                data: {
                    labels: classLabels,
                    datasets: [{
                        label: `Điểm TB (${milestone})`,
                        data: scoreData,
                        backgroundColor: '#3b82f6',
                        borderRadius: 8
                    }]
                },
                options: { 
                    responsive: true, 
                    maintainAspectRatio: false, 
                    plugins: { legend: { display: false } },
                    scales: {
                        x: { ticks: { font: { size: 13, weight: 'bold' } } },
                        y: { ticks: { font: { size: 12 } }, beginAtZero: true, max: 10 }
                    }
                }
            });

            let riskNameLabel = 'Nguy cơ cao (Y=0)';
            if(riskType === 'med') riskNameLabel = 'Nguy cơ thấp (Y=1)';
            else if(riskType === 'low') riskNameLabel = 'Không có nguy cơ (Y=2)';

            document.getElementById('chart-risk-title').innerText = `Biến Động Số Lượng Học Sinh (${riskNameLabel}) Qua Các Mốc T1 ➔ T4`;

            let lineDatasets = stats.map((st, idx) => {
                const colors = ['#f43f5e', '#3b82f6', '#10b981', '#f59e0b', '#8b5cf6'];
                return {
                    label: st.className,
                    data: st.trend,
                    borderColor: colors[idx % colors.length],
                    backgroundColor: colors[idx % colors.length],
                    tension: 0.3,
                    fill: false,
                    borderWidth: 3
                };
            });

            const ctxRisk = document.getElementById('classRiskComparisonChart').getContext('2d');
            if (classRiskChartInstance) classRiskChartInstance.destroy();
            classRiskChartInstance = new Chart(ctxRisk, {
                type: 'line',
                data: {
                    labels: ['T1 (Tuần 5)', 'T2 (Tuần 8)', 'T3 (Tuần 12)', 'T4 (Tuần 15)'],
                    datasets: lineDatasets
                },
                options: { 
                    responsive: true, 
                    maintainAspectRatio: false, 
                    plugins: { 
                        legend: { display: true, position: 'top' } 
                    },
                    scales: {
                        x: { ticks: { font: { size: 12, weight: 'bold' } } },
                        y: { ticks: { font: { size: 12 }, precision: 0 }, beginAtZero: true }
                    }
                }
            });
        }

        function renderTable() {
            const searchVal = document.getElementById('search-student').value.toLowerCase();
            const classVal = document.getElementById('filter-class').value;
            const riskVal = document.getElementById('filter-risk').value;

            const filtered = studentData.filter(s => {
                return s.id.toLowerCase().includes(searchVal) && (!classVal || s.class === classVal) && (!riskVal || s.riskLevel === riskVal);
            });

            document.getElementById('filter-count').innerText = filtered.length;
            document.getElementById('total-count').innerText = studentData.length;
            const tbody = document.getElementById('student-table-body');
            tbody.innerHTML = '';

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" class="text-center py-8 text-slate-400">Chưa có dữ liệu học sinh hoặc không tìm thấy kết quả</td></tr>`;
                return;
            }

            filtered.forEach(s => {
                let badgeClass = s.y_label === 0 ? 'bg-rose-50 text-rose-700 border-rose-200 font-bold' : (s.y_label === 1 ? 'bg-amber-50 text-amber-700 border-amber-200' : 'bg-emerald-50 text-emerald-700 border-emerald-200');

                tbody.innerHTML += `
                    <tr class="hover:bg-slate-50 transition">
                        <td class="py-3 px-4 font-bold text-slate-800">${s.id}</td>
                        <td class="py-3 px-4">${s.class}</td>
                        <td class="py-3 px-4 text-center font-semibold">${s.t1}đ | ${s.prob_t1}%</td>
                        <td class="py-3 px-4 text-center font-semibold">${s.t2}đ | ${s.prob_t2}%</td>
                        <td class="py-3 px-4 text-center font-semibold">${s.t3}đ | ${s.prob_t3}%</td>
                        <td class="py-3 px-4 text-center font-bold text-blue-600">${s.t4}đ | ${s.probability}%</td>
                        <td class="py-3 px-4 text-center"><span class="px-2.5 py-1 rounded-full text-[11px] border ${badgeClass}">Y = ${s.y_label} (${s.riskLevel})</span></td>
                        <td class="py-3 px-4 text-right"><button onclick="openModal('${s.id}')" class="px-2.5 py-1 bg-slate-100 hover:bg-slate-200 rounded-lg text-xs font-medium">Xem Hồ Sơ</button></td>
                    </tr>
                `;
            });
        }

        function filterStudents() { renderTable(); }

        function openModal(studentId) {
            const student = studentData.find(s => s.id === studentId);
            if (!student) return;
            currentSelectedStudent = student;
            
            document.getElementById('modal-student-id').innerText = `Mã HS: ${student.id}`;
            document.getElementById('modal-sub-info').innerText = `Lớp ${student.class} • Mốc hiện tại: T4`;
            document.getElementById('modal-risk-badge').innerText = student.riskLevel;
            document.getElementById('modal-prob').innerText = `${student.probability}%`;
            
            let trendText = 'Tiến độ ổn định';
            if (student.t4 < 6.5) trendText = 'Cần hỗ trợ cải thiện';
            else if (student.t4 >= 8.0) trendText = 'Phát triển rất tốt';
            document.getElementById('modal-trend-overall').innerText = trendText;

            const tbodyHistory = document.getElementById('modal-history-tbody');
            tbodyHistory.innerHTML = '';
            
            const history = [
                { milestone: 'T1 (Tuần 5 - Baseline)', a: student.t1, l6: Math.round(student.l6_1*100), prob: student.prob_t1, diff: 'Khởi điểm' },
                { milestone: 'T2 (Tuần 8)', a: student.t2, l6: Math.round(student.l6_2*100), prob: student.prob_t2, diff: (student.t2 - student.t1) },
                { milestone: 'T3 (Tuần 12)', a: student.t3, l6: Math.round(student.l6_3*100), prob: student.prob_t3, diff: (student.t3 - student.t2) },
                { milestone: 'T4 (Tuần 15)', a: student.t4, l6: Math.round(student.l6_4*100), prob: student.probability, diff: (student.t4 - student.t3) }
            ];

            history.forEach((h, idx) => {
                let diffStr = '';
                if (idx === 0) {
                    diffStr = `<span class="text-slate-400">Khởi điểm</span>`;
                } else {
                    let d = h.diff;
                    if (d > 0) diffStr = `<span class="text-emerald-600 font-bold">+${d.toFixed(2)}đ (Tăng điểm)</span>`;
                    else if (d < 0) diffStr = `<span class="text-rose-600 font-bold">${d.toFixed(2)}đ (Giảm điểm)</span>`;
                    else diffStr = `<span class="text-slate-500 font-semibold">0đ (Không đổi)</span>`;
                }

                tbodyHistory.innerHTML += `
                    <tr class="hover:bg-slate-50 transition">
                        <td class="p-3 font-semibold text-slate-700">${h.milestone}</td>
                        <td class="p-3 text-center font-bold text-slate-800">${h.a}đ</td>
                        <td class="p-3 text-center font-semibold text-blue-600">${h.l6}%</td>
                        <td class="p-3 text-center font-bold text-amber-700">${h.prob}%</td>
                        <td class="p-3 text-right font-medium">${diffStr}</td>
                    </tr>
                `;
            });

            // Cập nhật khối đóng góp A & L
            const lmsPct = Math.round(student.l6_4 * 100);
            document.getElementById('contrib-a-title').innerText = `Điểm bài kiểm tra (Group A: ${student.t4} điểm)`;
            document.getElementById('contrib-a-bar').style.width = `${Math.min(100, Math.max(10, student.t4 * 10))}%`;
            if (student.t4 < 6.5) {
                document.getElementById('contrib-a-status').innerText = 'Điểm dưới 6,5';
                document.getElementById('contrib-a-status').className = 'text-rose-600 font-bold';
                document.getElementById('contrib-a-bar').className = 'bg-rose-500 h-full rounded-full transition-all duration-500';
            } else {
                document.getElementById('contrib-a-status').innerText = 'Ổn định / Tích cực';
                document.getElementById('contrib-a-status').className = 'text-emerald-600 font-bold';
                document.getElementById('contrib-a-bar').className = 'bg-emerald-500 h-full rounded-full transition-all duration-500';
            }

            document.getElementById('contrib-l-title').innerText = `Tương tác LMS (Group L: ${lmsPct}%)`;
            document.getElementById('contrib-l-bar').style.width = `${lmsPct}%`;
            if (lmsPct < 50) {
                document.getElementById('contrib-l-status').innerText = 'Ít tương tác';
                document.getElementById('contrib-l-status').className = 'text-amber-600 font-bold';
            } else {
                document.getElementById('contrib-l-status').innerText = 'Tích cực';
                document.getElementById('contrib-l-status').className = 'text-blue-600 font-bold';
            }

            document.getElementById('plan-milestone').value = '4';
            renderLearningPlan();
            document.getElementById('student-modal').classList.remove('hidden');
        }

        // Use textContent for recommendation text, including imported student identifiers.
        function selectedLearningPlan() {
            const m = document.getElementById('plan-milestone').value;
            return currentSelectedStudent?.learning_plans?.['T' + m];
        }
        function learningPlanText(plan) {
            return [plan.title, 'Căn cứ: ' + plan.reasons.join('; '),
                'Việc cần làm: ' + plan.actions.join('; '), 'Lịch học: ' + plan.schedule,
                'Mục tiêu: ' + plan.target, 'Đánh giá lại: sau ' + plan.days + ' ngày kể từ khi áp dụng.',
                ...plan.notes].join(String.fromCharCode(10));
        }
        function renderLearningPlan() {
            const plan = selectedLearningPlan();
            document.getElementById('learning-plan-content').textContent = plan
                ? learningPlanText(plan) : 'Chưa có đủ dữ liệu để tạo phương án. Vui lòng tải lại tệp dữ liệu.';
        }
        function applyLearningPlan() {
            const plan = selectedLearningPlan();
            if (!plan || !currentSelectedStudent) return;
            const note = document.getElementById('form-note');
            if (note.value.trim() && !confirm('Thay nội dung ghi chú đang soạn bằng phương án này?')) return;
            navigateInterventionFromModal();
            document.getElementById('form-action').value = 'Phương án học tập cá nhân';
            note.value = learningPlanText(plan);
            const due = new Date(); due.setDate(due.getDate() + plan.days);
            document.getElementById('form-reassess-date').value = [due.getFullYear(),
                String(due.getMonth()+1).padStart(2, '0'), String(due.getDate()).padStart(2, '0')].join('-');
            document.getElementById('form-status').value = 'Chưa can thiệp';
        }

        function closeModal() { document.getElementById('student-modal').classList.add('hidden'); }
        function navigateInterventionFromModal() { closeModal(); switchTab('tab-interventions'); document.getElementById('form-student-id').value = currentSelectedStudent.id; }

        function populateInterventionStudentOptions() {
            const select = document.getElementById('form-student-id');
            const classFilter = document.getElementById('filter-class');
            if(!select) return;
            select.innerHTML = '';
            
            const classes = [...new Set(studentData.map(s => s.class))];
            if(classFilter) {
                classFilter.innerHTML = '<option value="">Tất cả các lớp</option>';
                classes.forEach(c => { classFilter.innerHTML += `<option value="${c}">${c}</option>`; });
            }

            studentData.forEach(s => { select.innerHTML += `<option value="${s.id}">${s.id} (${s.class})</option>`; });
        }

        function renderInterventions() {
            const container = document.getElementById('interventions-list');
            if(!container) return;
            document.getElementById('intervention-count').innerText = interventionLogs.length;
            container.innerHTML = '';
            if (interventionLogs.length === 0) {
                container.innerHTML = '<div class="text-center py-8 text-slate-400 text-xs">Chưa có nhật ký can thiệp nào</div>';
                return;
            }
            interventionLogs.forEach(log => {
                container.innerHTML += `
                    <div class="p-3.5 bg-slate-50 rounded-xl border border-slate-200 space-y-1.5 text-xs">
                        <div class="flex justify-between font-bold"><span>${log.studentId} • ${log.date}</span><span class="px-2 py-0.5 rounded border bg-white">${log.status}</span></div>
                        <div><strong>Biện pháp:</strong> ${log.action}</div>
                        <div class="italic text-slate-500">"${log.note}"</div>
                    </div>
                `;
            });
        }

        function saveIntervention(e) {
            e.preventDefault();
            interventionLogs.unshift({
                id: Date.now(),
                studentId: document.getElementById('form-student-id').value,
                date: new Date().toISOString().split('T')[0],
                action: document.getElementById('form-action').value,
                note: document.getElementById('form-note').value || 'Không có ghi chú',
                reassessDate: document.getElementById('form-reassess-date').value,
                status: document.getElementById('form-status').value
            });
            renderInterventions();
            alert('Lưu nhật ký can thiệp thành công!');
            document.getElementById('form-note').value = '';
        }

        function handleExcelUpload(e) {
            const file = e.target.files && e.target.files[0];
            if (!file) return;
            const status = document.getElementById('upload-status');
            status.textContent = `Đang tải lên ${file.name}...`;
            const reader = new FileReader();
            reader.onload = () => {
                const bytes = new Uint8Array(reader.result);
                let binary = '';
                const chunkSize = 0x8000;
                for (let i = 0; i < bytes.length; i += chunkSize) {
                    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
                }
                const fileBase64 = btoa(binary);
                window.parent.postMessage({
                    isStreamlitMessage: true,
                    type: 'streamlit:setComponentValue',
                    value: { file_name: file.name, file_base64: fileBase64 },
                    dataType: 'json'
                }, '*');
                status.textContent = 'Đã gửi tệp lên hệ thống. Đang xử lý...';
            };
            reader.onerror = () => {
                status.textContent = 'Không thể đọc tệp. Vui lòng thử lại.';
            };
            reader.readAsArrayBuffer(file);
        }

        function exportExcel() {
            if (studentData.length === 0) {
                alert('Chưa có dữ liệu học sinh để xuất Excel!');
                return;
            }
            
            const mainRows = studentData.map(s => ({
                'Student_ID': s.id,
                'Class_ID': s.class,
                'TX1': s.tx1,
                'TX2': s.tx2,
                'GK1': s.gk1,
                'TX3': s.tx3,
                'T1 (Điểm TB)': s.t1,
                'T2 (Điểm TB)': s.t2,
                'T3 (Điểm TB)': s.t3,
                'T4 (Điểm TB)': s.t4,
                'Xác_Suất_Nguy_Co (%)': s.probability,
                'Mức_Cảnh_Báo (Y)': s.y_label,
                'Nhãn_Mức_Độ': s.riskLevel
            }));

            const featureRows = studentData.map(s => ({
                'Student_ID': s.id,
                'Class_ID': s.class,
                'A1_Diem_TB': s.f_t4.a1,
                'A2_Diem_max': s.f_t4.a2,
                'A3_Diem_min': s.f_t4.a3,
                'A4_Khoang_diem': s.f_t4.a4,
                'A5_Do_lech_chuan_mau': s.f_t4.a5,
                'A6_Xu_huong_OLS': s.f_t4.a6,
                'A7_Diem_gan_nhat': s.f_t4.a7,
                'L1_Dang_nhap': s.f_t4.l1,
                'L2_Phien_hoc': s.f_t4.l2,
                'L3_Thoi_gian_phut': s.f_t4.l3,
                'L4_Xem_hoc_lieu': s.f_t4.l4,
                'L5_Lam_quiz': s.f_t4.l5,
                'L6_Ty_le_tuan_da_nop': s.f_t4.l6,
                'Y': s.y_label
            }));

            const wb = XLSX.utils.book_new();
            const wsMain = XLSX.utils.json_to_sheet(mainRows);
            const wsFeatures = XLSX.utils.json_to_sheet(featureRows);

            XLSX.utils.book_append_sheet(wb, wsMain, "Ket_Qua_Chinh");
            XLSX.utils.book_append_sheet(wb, wsFeatures, "Bang_13_Dac_Trung");

            const planRows = studentData.flatMap(s => Object.entries(s.learning_plans || {}).map(([m, p]) => ({
                Student_ID: s.id, Class_ID: s.class, Moc: m, Nhom_du_bao: p.label,
                Phuong_an: p.title, Can_cu: p.reasons.join('; '), Viec_can_lam: p.actions.join('; '),
                Lich_hoc: p.schedule, Muc_tieu: p.target, Danh_gia_sau_ngay: p.days,
                Luu_y: p.notes.join('; ')
            })));
            XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(planRows), 'Phuong_An_Hoc_Tap');
            XLSX.writeFile(wb, "Bao_Cao_Canh_Bao_Tin_Hoc_12_3_Bang.xlsx");
        }
    </script>
</body>
</html>
"""

# ================= RANDOM FOREST: HUẤN LUYỆN ĐỘC LẬP CHO TỪNG MỐC T1 -> T4 =================
FEATURE_COLS = [
    'A1_Diem_TB', 'A2_Diem_max', 'A3_Diem_min', 'A4_Khoang_diem',
    'A5_Do_lech_chuan_mau', 'A6_Xu_huong_OLS', 'A7_Diem_gan_nhat',
    'L1_Dang_nhap', 'L2_Phien_hoc', 'L3_Thoi_gian_phut',
    'L4_Xem_hoc_lieu', 'L5_Lam_quiz', 'L6_Ty_le_tuan_da_nop'
]

def _number(row, names, default):
    for name in names:
        value = row.get(name, None)
        if pd.notna(value):
            try:
                return float(value)
            except (TypeError, ValueError):
                pass
    return float(default)

def _features(grades, logins, minutes, tasks, max_tasks):
    grades = np.asarray(grades, dtype=float)
    n = len(grades)
    x = np.arange(1, n + 1, dtype=float)
    slope = float(np.polyfit(x, grades, 1)[0]) if n > 1 else 0.0
    mean = float(np.mean(grades))
    return {
        'A1_Diem_TB': mean,
        'A2_Diem_max': float(np.max(grades)),
        'A3_Diem_min': float(np.min(grades)),
        'A4_Khoang_diem': float(np.max(grades) - np.min(grades)),
        'A5_Do_lech_chuan_mau': float(np.std(grades, ddof=0)),
        'A6_Xu_huong_OLS': slope,
        'A7_Diem_gan_nhat': float(grades[-1]),
        'L1_Dang_nhap': float(logins),
        'L2_Phien_hoc': float(max(1, round(logins / 2))),
        'L3_Thoi_gian_phut': float(minutes),
        'L4_Xem_hoc_lieu': float(round(logins * 1.5)),
        'L5_Lam_quiz': float(round(tasks * 0.8)),
        'L6_Ty_le_tuan_da_nop': float(min(1.0, tasks / max_tasks)) if max_tasks else 0.0,
    }

def _predict_one(model, features):
    row = pd.DataFrame([[features[c] for c in FEATURE_COLS]], columns=FEATURE_COLS)
    probabilities = model.predict_proba(row)[0]
    classes = list(model.classes_)
    predicted = int(classes[int(np.argmax(probabilities))])
    risk_probability = float(probabilities[classes.index(0)]) if 0 in classes else 0.0
    return predicted, round(risk_probability * 100)

# ================= ĐỀ XUẤT HỌC TẬP THEO QUY TẮC (KHÔNG PHẢI SHAP) =================
def _observed_number(row, names, minimum=0, maximum=None):
    """Only use real, finite inputs as evidence for learning recommendations."""
    for name in names:
        try:
            value = float(row.get(name))
            if np.isfinite(value) and value >= minimum and (maximum is None or value <= maximum):
                return value
        except (TypeError, ValueError):
            continue
    return None


def _learning_plan(row, milestone, predicted):
    labels = ['Nguy cơ cao', 'Nguy cơ thấp', 'Không có nguy cơ']
    names = [('TX1', 'tx1'), ('TX2', 'tx2'), ('GK1', 'gk1'), ('TX3', 'tx3')][:milestone]
    values = [_observed_number(row, keys, maximum=10) for keys in names]
    complete = all(v is not None for v in values)
    tasks = _observed_number(row, [f'Bai_hoan_thanh_T{milestone}'])
    planned = [3, 6, 10, 15][milestone - 1]
    reasons, actions, notes = [], [], []
    mean = float(np.mean(values)) if complete else None
    declining = complete and len(values) > 1 and float(np.polyfit(range(len(values)), values, 1)[0]) < -0.2

    if mean is not None:
        reasons.append(f'Điểm TB T{milestone}: {mean:.1f}/10')
        if mean < 6.5:
            actions.append('Ôn khái niệm, làm 5 bài cơ bản')
        elif mean < 8:
            actions.append('Chữa lỗi, làm 3 bài cơ bản + 2 vận dụng')
        else:
            actions.append('Duy trì ôn tập, làm 1 dự án nhỏ')
    else:
        notes.append('Thiếu điểm bài tập.')
        actions.append('GV kiểm tra bài gần nhất')

    if declining:
        reasons.append('Điểm đang giảm')
        actions.append('Gặp GV chữa lỗi, làm lại bài')

    if tasks is not None and tasks <= planned:
        reasons.append(f'Làm {tasks:g}/{planned} bài')
        if tasks / planned < 0.8:
            actions.append('Nộp bù 2 bài thiếu, nhắc lịch học')

    urgent = (mean is not None and mean < 6.5) or declining or predicted == 0
    days = 7 if urgent else 14
    schedule = '4 buổi x 25p/tuần' if urgent else '3 buổi x 30p/tuần'
    
    target = 'Hoàn thành bài giao'
    if mean is not None:
        target += f', đạt {"6.5" if mean < 6.5 else "8.0" if mean < 8 else ">8.0"}'

    notes.append('GV tự điều chỉnh nội dung.')

    return dict(title=f'Kế hoạch T{milestone} — {labels[predicted]}',
                label=labels[predicted], reasons=reasons or ['Chưa đủ dữ liệu'],
                actions=actions, schedule=schedule, target=target, days=days, notes=notes)


@st.cache_data(show_spinner=False)
def cached_build_predictions(file_bytes, file_name):
    file_buffer = BytesIO(file_bytes)
    if file_name.lower().endswith('.csv'):
        source_df = pd.read_csv(file_buffer)
    else:
        source_df = pd.read_excel(file_buffer)

    rows = []
    for position, (_, row) in enumerate(source_df.iterrows(), start=1):
        sid = str(row.get('Student_ID', row.get('Mã Học Sinh', row.get('Mã HS', f'HS12_{position:03d}')))).strip()
        cls = str(row.get('Class_ID', row.get('Lớp', '12A1'))).strip()
        tx1 = _number(row, ['TX1', 'tx1'], 7.0)
        tx2 = _number(row, ['TX2', 'tx2'], 7.0)
        gk1 = _number(row, ['GK1', 'gk1'], 7.0)
        tx3 = _number(row, ['TX3', 'tx3'], 7.0)
        
        grades_by_term = [[tx1], [tx1, tx2], [tx1, tx2, gk1], [tx1, tx2, gk1, tx3]]
        logins = [_number(row, [f'Dang_nhap_T{i}'], default) for i, default in enumerate([5, 12, 20, 28], 1)]
        minutes = [_number(row, [f'Tgian_online_T{i} (phút)', f'Tgian_online_T{i}'], default) for i, default in enumerate([30, 80, 150, 220], 1)]
        tasks = [_number(row, [f'Bai_hoan_thanh_T{i}'], default) for i, default in enumerate([2, 5, 8, 12], 1)]
        max_tasks = [3, 6, 10, 15]
        
        feats = [_features(grades_by_term[i], logins[i], minutes[i], tasks[i], max_tasks[i]) for i in range(4)]
        t1, t2, t3, t4 = [f['A1_Diem_TB'] for f in feats]
        
        rows.append({
            'id': sid, 'class': cls, 'tx1': tx1, 'tx2': tx2, 'gk1': gk1, 'tx3': tx3,
            't1': round(t1, 2), 't2': round(t2, 2), 't3': round(t3, 2), 't4': round(t4, 2),
            'features': feats
        })

    if not rows:
        raise ValueError('Tệp không có dữ liệu học sinh.')

    X_t1 = pd.DataFrame([r['features'][0] for r in rows], columns=FEATURE_COLS)
    X_t2 = pd.DataFrame([r['features'][1] for r in rows], columns=FEATURE_COLS)
    X_t3 = pd.DataFrame([r['features'][2] for r in rows], columns=FEATURE_COLS)
    X_t4 = pd.DataFrame([r['features'][3] for r in rows], columns=FEATURE_COLS)

    if 'Y' in source_df.columns:
        y = pd.to_numeric(source_df['Y'], errors='coerce').fillna(-1).astype(int).to_numpy()
        if not np.isin(y, [0, 1, 2]).all():
            raise ValueError('Cột Y chỉ được chứa các nhãn 0, 1 hoặc 2.')
        label_source = 'nhãn Y có sẵn trong tệp'
    else:
        scores = X_t4['A1_Diem_TB'].to_numpy()
        y = np.where(scores < 6.5, 0, np.where(scores < 8.0, 1, 2))
        label_source = 'nhãn tham chiếu tạo từ ngưỡng điểm (<6,5; 6,5–8; >=8)'

    model_t1 = RandomForestClassifier(n_estimators=100, max_depth=6, class_weight='balanced', random_state=42, n_jobs=-1)
    model_t2 = RandomForestClassifier(n_estimators=100, max_depth=6, class_weight='balanced', random_state=42, n_jobs=-1)
    model_t3 = RandomForestClassifier(n_estimators=100, max_depth=6, class_weight='balanced', random_state=42, n_jobs=-1)
    model_t4 = RandomForestClassifier(n_estimators=100, max_depth=6, class_weight='balanced', random_state=42, n_jobs=-1)

    model_t1.fit(X_t1, y)
    model_t2.fit(X_t2, y)
    model_t3.fit(X_t3, y)
    model_t4.fit(X_t4, y)

    def to_js_features(f):
        return {
            'a1': f['A1_Diem_TB'], 'a2': f['A2_Diem_max'], 'a3': f['A3_Diem_min'],
            'a4': f['A4_Khoang_diem'], 'a5': f['A5_Do_lech_chuan_mau'],
            'a6': f['A6_Xu_huong_OLS'], 'a7': f['A7_Diem_gan_nhat'],
            'l1': f['L1_Dang_nhap'], 'l2': f['L2_Phien_hoc'], 'l3': f['L3_Thoi_gian_phut'],
            'l4': f['L4_Xem_hoc_lieu'], 'l5': f['L5_Lam_quiz'], 'l6': f['L6_Ty_le_tuan_da_nop']
        }

    for r, (_, source_row) in zip(rows, source_df.iterrows()):
        p1 = _predict_one(model_t1, r['features'][0])
        p2 = _predict_one(model_t2, r['features'][1])
        p3 = _predict_one(model_t3, r['features'][2])
        p4 = _predict_one(model_t4, r['features'][3])

        labels = [p1[0], p2[0], p3[0], p4[0]]
        risks = [p1[1], p2[1], p3[1], p4[1]]
        final_label = labels[3]

        r['learning_plans'] = {f'T{i}': _learning_plan(source_row, i, labels[i-1]) for i in range(1, 5)}
        r.update({
            'prob_t1': risks[0], 'prob_t2': risks[1], 'prob_t3': risks[2],
            'probability': risks[3], 'y_label': final_label,
            'y_label_t1': labels[0], 'y_label_t2': labels[1], 'y_label_t3': labels[2],
            'riskLevel': ['Nguy cơ cao', 'Nguy cơ thấp', 'Không có nguy cơ'][final_label],
            'f_t1': to_js_features(r['features'][0]), 'f_t2': to_js_features(r['features'][1]),
            'f_t3': to_js_features(r['features'][2]), 'f_t4': to_js_features(r['features'][3]),
            'a6_1': r['features'][0]['A6_Xu_huong_OLS'],
            'a6_2': r['features'][1]['A6_Xu_huong_OLS'],
            'a6_3': r['features'][2]['A6_Xu_huong_OLS'],
            'a6_4': r['features'][3]['A6_Xu_huong_OLS'],
            'l6_1': r['features'][0]['L6_Ty_le_tuan_da_nop'],
            'l6_2': r['features'][1]['L6_Ty_le_tuan_da_nop'],
            'l6_3': r['features'][2]['L6_Ty_le_tuan_da_nop'],
            'l6_4': r['features'][3]['L6_Ty_le_tuan_da_nop'],
        })
        del r['features']
        
    return rows, label_source

# ================= STREAMLIT COMPONENT & BRIDGE =================
component_dir = Path(__file__).parent / ".edurisk_dashboard_component"
component_dir.mkdir(exist_ok=True)
component_html = edurisk_html_code.replace('__RF_STUDENT_DATA__', '[]')

component_bridge = r"""
<script>
(function () {
    window.parent.postMessage({
        isStreamlitMessage: true,
        type: 'streamlit:componentReady',
        apiVersion: 1
    }, '*');

    window.addEventListener('message', function (event) {
        const msg = event.data || {};
        if (!msg || msg.type !== 'streamlit:render') return;
        const args = msg.args || {};
        if (Array.isArray(args.student_data)) {
            studentData = args.student_data;
            if (document.getElementById('main-app') &&
                !document.getElementById('main-app').classList.contains('opacity-0')) {
                renderDashboard();
                renderTable();
                renderComparisonTable();
                populateInterventionStudentOptions();
            }
        }
        const statusEl = document.getElementById('upload-status');
        if (statusEl && args.upload_status) statusEl.textContent = args.upload_status;
    });

    document.addEventListener('change', function (event) {
        if (!event.target || event.target.id !== 'excel-upload-input') return;
        handleExcelUpload(event);
    });
    window.parent.postMessage({
        isStreamlitMessage: true,
        type: 'streamlit:setFrameHeight',
        height: 1200
    }, '*');
})();
</script>
"""
component_html = component_html.replace('</body>', component_bridge + '\n</body>')
(component_dir / 'index.html').write_text(component_html, encoding='utf-8')

dashboard_component = components.declare_component(
    'edurisk_dashboard_with_upload_v3',
    path=str(component_dir)
)

if 'rf_student_data' not in st.session_state:
    st.session_state.rf_student_data = []
if 'rf_upload_hash' not in st.session_state:
    st.session_state.rf_upload_hash = ''
if 'rf_upload_status' not in st.session_state:
    st.session_state.rf_upload_status = ''

upload_payload = dashboard_component(
    student_data=st.session_state.rf_student_data,
    upload_status=st.session_state.rf_upload_status,
    default=None,
    key='edurisk_dashboard_v3'
)

if isinstance(upload_payload, dict) and upload_payload.get('file_base64'):
    encoded_file = upload_payload['file_base64']
    upload_hash = hashlib.sha256(encoded_file.encode('utf-8')).hexdigest()
    if upload_hash != st.session_state.rf_upload_hash:
        try:
            raw_file = base64.b64decode(encoded_file)
            file_name = str(upload_payload.get('file_name', 'students.xlsx'))
            rf_student_data, rf_label_source = cached_build_predictions(raw_file, file_name)
            st.session_state.rf_student_data = rf_student_data
            st.session_state.rf_upload_hash = upload_hash
            st.session_state.rf_upload_status = (
                f'✓ Tải thành công {file_name} • Đã dự đoán {len(rf_student_data)} học sinh.'
            )
        except Exception as exc:
            st.session_state.rf_upload_hash = upload_hash
            st.session_state.rf_upload_status = f'Không thể xử lý tệp: {exc}'
        st.rerun()

# streamlit run app.py