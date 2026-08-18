/* =========================================================================
   AWS CLOUD COST ESTIMATOR & MIGRATION ADVISOR - LOGIC ENGINE
   ========================================================================= */

document.addEventListener('DOMContentLoaded', () => {
    // 1. Initialize Lucide Icons
    lucide.createIcons();

    // 2. DOM Elements Selection
    const elDropzone = document.getElementById('excel-dropzone');
    const elFileInput = document.getElementById('excel-file-input');
    const elFileBadge = document.getElementById('loaded-file-badge');
    const elFileName = document.getElementById('loaded-file-name');
    const elBtnRemoveFile = document.getElementById('btn-remove-file');

    const elAwsRegion = document.getElementById('aws-region');
    const elAwsPlan = document.getElementById('aws-pricing-plan');
    const elAncillaryPct = document.getElementById('input-ancillary-pct');
    const elValAncillaryPct = document.getElementById('val-ancillary-pct');

    const elStatTotalApps = document.getElementById('stat-total-apps');
    const elStatTotalComponents = document.getElementById('stat-total-components');
    const elStatOnPremCost = document.getElementById('stat-onprem-cost');
    const elStatAwsCost = document.getElementById('stat-aws-cost');
    const elStatAwsCostIcon = document.getElementById('stat-aws-cost-icon');

    const elTableBody = document.getElementById('results-table-body');
    const elSearchInput = document.getElementById('search-input');
    const elBtnExportExcel = document.getElementById('btn-export-excel');
    const elBtnPrintReport = document.getElementById('btn-print-report');

    const elBtnDownloadTemplate = document.getElementById('btn-download-template');
    const elBtnLoadSample = document.getElementById('btn-load-sample');

    // State Variables
    let parsedPortfolio = []; // Holds the raw parsed apps and components
    let processedPortfolio = []; // Holds calculated cost data
    let onPremVsAwsChart = null;
    let awsBreakdownChart = null;

    // Sorting state variables
    let currentSortColumn = 'stt';
    let currentSortDirection = 'asc';
    let sortPortfolio;
    let updateHeaderSortUI;

    // Group mappings and properties
    const DEFAULT_GROUP_INFO = {
        "N1": {
            code: "N1",
            name: "N1 — Ứng dụng web, quản trị nội bộ",
            stability: "Mức 4 (Rất ổn định)",
            performance: "Mức 3 (Đáp ứng cao hơn yêu cầu nghiệp vụ.)",
            stabilityAbbr: "M4",
            performanceAbbr: "M3",
            dbChangeRate: 5
        },
        "N2": {
            code: "N2",
            name: "N2 — Kênh khách hàng, tác nghiệp phụ thuộc Core",
            stability: "Mức 2 (Cơ bản ổn định)",
            performance: "Mức 1 (Không đáp ứng yêu cầu nghiệp vụ)",
            stabilityAbbr: "M2",
            performanceAbbr: "M1",
            dbChangeRate: 10
        },
        "N3": {
            code: "N3",
            name: "N3 — Thanh toán, thẻ, chứng khoán, ký số",
            stability: "Mức 1 (Không ổn định)",
            performance: "Mức 1 (Không đáp ứng yêu cầu nghiệp vụ)",
            stabilityAbbr: "M1",
            performanceAbbr: "M1",
            dbChangeRate: 15
        },
        "N4": {
            code: "N4",
            name: "N4 — Dữ liệu, báo cáo, xử lý theo lô",
            stability: "Mức 4 nếu dữ liệu đi cùng; mức 1 nếu dữ liệu ở lại on-prem",
            performance: "Mức 3 nếu dữ liệu đi cùng; mức 1 nếu dữ liệu ở lại on-prem",
            stabilityAbbr: "M4/M1",
            performanceAbbr: "M3/M1",
            dbChangeRate: 20
        },
        "N5": {
            code: "N5",
            name: "N5 — Nền tảng xác thực, tích hợp dùng chung",
            stability: "Mức 2 nếu chuyển trước các ứng dụng; mức 4 nếu chuyển sau phần lớn ứng dụng",
            performance: "Mức 1 (Không đáp ứng yêu cầu nghiệp vụ) nếu chuyển trước; mức 3 nếu chuyển sau",
            stabilityAbbr: "M2/M4",
            performanceAbbr: "M1/M3",
            dbChangeRate: 8
        },
        "N6": {
            code: "N6",
            name: "N6 — Phần mềm thương mại đóng gói",
            stability: "Mức 3 nếu được hãng chứng nhận hỗ trợ; mức 1 (Không đáp ứng yêu cầu nghiệp vụ) nếu không được hỗ trợ",
            performance: "Mức 2 (Cơ bản ổn định)",
            stabilityAbbr: "M3/M1",
            performanceAbbr: "M2",
            dbChangeRate: 5
        },
        "N7": {
            code: "N7",
            name: "N7 — AI và GenAI",
            stability: "Mức 4 (Rất ổn định)",
            performance: "Mức 3 (Đáp ứng cao hơn yêu cầu nghiệp vụ.)",
            stabilityAbbr: "M4",
            performanceAbbr: "M3",
            dbChangeRate: 12
        }
    };

    let GROUP_INFO = JSON.parse(JSON.stringify(DEFAULT_GROUP_INFO));

    const DEFAULT_COMPLEXITY_INFO = [
        { level: "Mức 1", minComp: 1, maxComp: 5, mandays: 135, note: "Độ phức tạp thấp nhất" },
        { level: "Mức 2", minComp: 6, maxComp: 9, mandays: 270, note: "Độ phức tạp trung bình thấp" },
        { level: "Mức 3", minComp: 10, maxComp: 19, mandays: 400, note: "Độ phức tạp trung bình" },
        { level: "Mức 4", minComp: 20, maxComp: 39, mandays: 800, note: "Độ phức tạp trung bình cao" },
        { level: "Mức 5", minComp: 40, maxComp: 9999, mandays: 1000, note: "Độ phức tạp cao nhất" }
    ];

    const DEFAULT_MANDAY_RATE = 2000000;

    let COMPLEXITY_INFO = [];
    let MANDAY_RATE = DEFAULT_MANDAY_RATE;

    const normalizeGroup = (groupStr) => {
        if (!groupStr) return "N1";
        const clean = String(groupStr).trim().toUpperCase();
        const match = clean.match(/^(N[1-7])/i);
        if (match) {
            return match[1];
        }
        return "N1";
    };

    // 3. DEFAULT STATIC AWS PRICING CATALOG (Fallback reference)
    const DEFAULT_INSTANCE_CATALOG = {
        ec2: [
            { name: 't3a.small', cpu: 2, ram: 2, rates: { 'us-east-1': 22/730, 'ap-southeast-1': 22/730, 'ap-northeast-1': 22/730, 'eu-central-1': 22/730 } },
            { name: 'c6a.large', cpu: 2, ram: 4, rates: { 'us-east-1': 67/730, 'ap-southeast-1': 67/730, 'ap-northeast-1': 67/730, 'eu-central-1': 67/730 } },
            { name: 'm6i.large', cpu: 2, ram: 8, rates: { 'us-east-1': 87.6/730, 'ap-southeast-1': 87.6/730, 'ap-northeast-1': 87.6/730, 'eu-central-1': 87.6/730 } },
            { name: 'c6a.xlarge', cpu: 4, ram: 8, rates: { 'us-east-1': 135/730, 'ap-southeast-1': 135/730, 'ap-northeast-1': 135/730, 'eu-central-1': 135/730 } },
            { name: 'm6i.xlarge', cpu: 4, ram: 16, rates: { 'us-east-1': 180/730, 'ap-southeast-1': 180/730, 'ap-northeast-1': 180/730, 'eu-central-1': 180/730 } },
            { name: 'c6i.2xlarge', cpu: 8, ram: 16, rates: { 'us-east-1': 270/730, 'ap-southeast-1': 270/730, 'ap-northeast-1': 270/730, 'eu-central-1': 270/730 } },
            { name: 'm6i.2xlarge', cpu: 8, ram: 32, rates: { 'us-east-1': 350.4/730, 'ap-southeast-1': 350.4/730, 'ap-northeast-1': 350.4/730, 'eu-central-1': 350.4/730 } },
            { name: 'c6i.4xlarge', cpu: 16, ram: 32, rates: { 'us-east-1': 540/730, 'ap-southeast-1': 540/730, 'ap-northeast-1': 540/730, 'eu-central-1': 540/730 } },
            { name: 'c5a.4xlarge', cpu: 16, ram: 32, rates: { 'us-east-1': 540/730, 'ap-southeast-1': 540/730, 'ap-northeast-1': 540/730, 'eu-central-1': 540/730 } },
            { name: 'm6i.4xlarge', cpu: 16, ram: 64, rates: { 'us-east-1': 720/730, 'ap-southeast-1': 720/730, 'ap-northeast-1': 720/730, 'eu-central-1': 720/730 } },
            { name: 'c6i.8xlarge', cpu: 32, ram: 64, rates: { 'us-east-1': 1080/730, 'ap-southeast-1': 1080/730, 'ap-northeast-1': 1080/730, 'eu-central-1': 1080/730 } },
            { name: 'm6i.8xlarge', cpu: 32, ram: 128, rates: { 'us-east-1': 1401.6/730, 'ap-southeast-1': 1401.6/730, 'ap-northeast-1': 1401.6/730, 'eu-central-1': 1401.6/730 } }
        ],
        rds: [
            { name: 'db.m7g.large', cpu: 2, ram: 8, rates: { 'us-east-1': 395.52/730, 'ap-southeast-1': 395.52/730, 'ap-northeast-1': 395.52/730, 'eu-central-1': 395.52/730 } },
            { name: 'db.m7g.xlarge', cpu: 4, ram: 16, rates: { 'us-east-1': 748.11/730, 'ap-southeast-1': 748.11/730, 'ap-northeast-1': 748.11/730, 'eu-central-1': 748.11/730 } },
            { name: 'db.m7g.2xlarge', cpu: 8, ram: 32, rates: { 'us-east-1': 1500.74/730, 'ap-southeast-1': 1500.74/730, 'ap-northeast-1': 1500.74/730, 'eu-central-1': 1500.74/730 } },
            { name: 'db.m7g.4xlarge', cpu: 16, ram: 64, rates: { 'us-east-1': 2973.15/730, 'ap-southeast-1': 2973.15/730, 'ap-northeast-1': 2973.15/730, 'eu-central-1': 2973.15/730 } },
            { name: 'db.m7g.8xlarge', cpu: 32, ram: 128, rates: { 'us-east-1': 5918.70/730, 'ap-southeast-1': 5918.70/730, 'ap-northeast-1': 5918.70/730, 'eu-central-1': 5918.70/730 } },
            { name: 'db.m7g.16xlarge', cpu: 64, ram: 256, rates: { 'us-east-1': 11810.53/730, 'ap-southeast-1': 11810.53/730, 'ap-northeast-1': 11810.53/730, 'eu-central-1': 11810.53/730 } }
        ],
        msk: [
            { name: 'kafka.m7g.large', cpu: 2, ram: 8, rates: { 'us-east-1': 198.15/730, 'ap-southeast-1': 198.15/730, 'ap-northeast-1': 198.15/730, 'eu-central-1': 198.15/730 } },
            { name: 'kafka.m7g.xlarge', cpu: 4, ram: 16, rates: { 'us-east-1': 384.3/730, 'ap-southeast-1': 384.3/730, 'ap-northeast-1': 384.3/730, 'eu-central-1': 384.3/730 } },
            { name: 'kafka.m7g.2xlarge', cpu: 8, ram: 32, rates: { 'us-east-1': 756.6/730, 'ap-southeast-1': 756.6/730, 'ap-northeast-1': 756.6/730, 'eu-central-1': 756.6/730 } },
            { name: 'kafka.m7g.4xlarge', cpu: 16, ram: 64, rates: { 'us-east-1': 1501.2/730, 'ap-southeast-1': 1501.2/730, 'ap-northeast-1': 1501.2/730, 'eu-central-1': 1501.2/730 } },
            { name: 'kafka.m7g.8xlarge', cpu: 32, ram: 128, rates: { 'us-east-1': 2990.4/730, 'ap-southeast-1': 2990.4/730, 'ap-northeast-1': 2990.4/730, 'eu-central-1': 2990.4/730 } }
        ],
        dms: [
            { name: 't3.small', cpu: 2, ram: 2, rates: { 'us-east-1': 95.56/730, 'ap-southeast-1': 95.56/730, 'ap-northeast-1': 95.56/730, 'eu-central-1': 95.56/730 }, note: 'Rất nhỏ / DB nhỏ, CDC thấp' },
            { name: 't3.medium', cpu: 2, ram: 4, rates: { 'us-east-1': 177.32/730, 'ap-southeast-1': 177.32/730, 'ap-northeast-1': 177.32/730, 'eu-central-1': 177.32/730 }, note: 'Nhỏ' },
            { name: 't3.large', cpu: 2, ram: 8, rates: { 'us-east-1': 340.84/730, 'ap-southeast-1': 340.84/730, 'ap-northeast-1': 340.84/730, 'eu-central-1': 340.84/730 }, note: 'Nhỏ–trung bình' },
            { name: 'c6i.large', cpu: 2, ram: 4, rates: { 'us-east-1': 338.58/730, 'ap-southeast-1': 338.58/730, 'ap-northeast-1': 338.58/730, 'eu-central-1': 338.58/730 }, note: 'Trung bình, CDC liên tục' },
            { name: 'c6i.xlarge', cpu: 4, ram: 8, rates: { 'us-east-1': 649.56/730, 'ap-southeast-1': 649.56/730, 'ap-northeast-1': 649.56/730, 'eu-central-1': 649.56/730 }, note: 'Trung bình–lớn' },
            { name: 'c6i.2xlarge', cpu: 8, ram: 16, rates: { 'us-east-1': 1270.06/730, 'ap-southeast-1': 1270.06/730, 'ap-northeast-1': 1270.06/730, 'eu-central-1': 1270.06/730 }, note: 'Lớn / CDC cao' },
            { name: 'r6i.large', cpu: 2, ram: 16, rates: { 'us-east-1': 458.30/730, 'ap-southeast-1': 458.30/730, 'ap-northeast-1': 458.30/730, 'eu-central-1': 458.30/730 }, note: 'Memory-intensive' },
            { name: 'r6i.xlarge', cpu: 4, ram: 32, rates: { 'us-east-1': 889.00/730, 'ap-southeast-1': 889.00/730, 'ap-northeast-1': 889.00/730, 'eu-central-1': 889.00/730 }, note: 'Memory-intensive lớn' },
            { name: 'r6i.2xlarge', cpu: 8, ram: 64, rates: { 'us-east-1': 1750.40/730, 'ap-southeast-1': 1750.40/730, 'ap-northeast-1': 1750.40/730, 'eu-central-1': 1750.40/730 }, note: 'Rất lớn' }
        ],
        storage: {
            ebs: { 'us-east-1': 0.096, 'ap-southeast-1': 0.096, 'ap-northeast-1': 0.096, 'eu-central-1': 0.096 }, // gp3 EBS
            rds: { 'us-east-1': 0.138, 'ap-southeast-1': 0.138, 'ap-northeast-1': 0.138, 'eu-central-1': 0.138 }, // gp3 RDS
            msk: { 'us-east-1': 0.10, 'ap-southeast-1': 0.10, 'ap-northeast-1': 0.10, 'eu-central-1': 0.10 },      // gp3 MSK Storage
            dms: { 'us-east-1': 0.096, 'ap-southeast-1': 0.096, 'ap-northeast-1': 0.096, 'eu-central-1': 0.096 }      // gp3 DMS Storage
        },
        eksClusterYear: 876,
        dataTransferOutRate: 0.12
    };

    // Active working catalog state (mutable and synchronized with localStorage)
    let INSTANCE_CATALOG = JSON.parse(JSON.stringify(DEFAULT_INSTANCE_CATALOG));

    // 3b. EC2 MULTI-DIMENSIONAL SPLITTING SOLVER (Application & Other Components)
    const solveEC2Splitting = (cpuReqRaw, ramReqRaw, region) => {
        const cpuReq = (isNaN(cpuReqRaw) || cpuReqRaw === null || cpuReqRaw === undefined) ? 0 : cpuReqRaw;
        const ramReq = (isNaN(ramReqRaw) || ramReqRaw === null || ramReqRaw === undefined) ? 0 : ramReqRaw;
        let remCpu = cpuReq;
        let remRam = ramReq;
        const selected = {}; // name -> count

        // Sort catalog descending by size (cpu first, then ram) to find the largest candidate blocks
        const catalog = [...INSTANCE_CATALOG.ec2].sort((a, b) => b.cpu - a.cpu || b.ram - a.ram);

        // Stage 1: Greedy fit large instance blocks that can fully sit within the remaining demands
        for (const inst of catalog) {
            if (inst.cpu <= 0 || inst.ram <= 0) continue;
            
            const count = Math.min(
                Math.floor(remCpu / inst.cpu),
                Math.floor(remRam / inst.ram)
            );

            if (count > 0) {
                selected[inst.name] = (selected[inst.name] || 0) + count;
                remCpu -= count * inst.cpu;
                remRam -= count * inst.ram;
            }
        }

        // Stage 2: Greedy coverage of any leftover CPU or RAM using performance/rate score
        while (remCpu > 0 || remRam > 0) {
            let bestInst = null;
            let bestScore = -Infinity;

            for (const inst of catalog) {
                const usefulCpu = Math.min(inst.cpu, Math.max(0, remCpu));
                const usefulRam = Math.min(inst.ram, Math.max(0, remRam));

                if (usefulCpu === 0 && usefulRam === 0) continue;

                const rate = inst.rates[region] || 0.0001;
                // Performance Score: Covered resources weighted per Dollar rate
                const score = (usefulCpu + usefulRam / 4) / rate;

                if (score > bestScore) {
                    bestScore = score;
                    bestInst = inst;
                }
            }

            if (!bestInst) break; // Emergency exit

            selected[bestInst.name] = (selected[bestInst.name] || 0) + 1;
            remCpu -= bestInst.cpu;
            remRam -= bestInst.ram;
        }

        return selected;
    };

    // 4. PRICE MATCHING ALGORITHM
    const matchInstance = (serviceType, cpuReqRaw, ramReqRaw, region) => {
        const cpuReq = (isNaN(cpuReqRaw) || cpuReqRaw === null || cpuReqRaw === undefined) ? 0 : cpuReqRaw;
        const ramReq = (isNaN(ramReqRaw) || ramReqRaw === null || ramReqRaw === undefined) ? 0 : ramReqRaw;
        if (serviceType === 'rds') {
            // Find an exact match first
            const exactMatch = INSTANCE_CATALOG.rds.find(inst => inst.cpu === cpuReq && inst.ram === ramReq);
            if (exactMatch) {
                return {
                    name: exactMatch.name,
                    rate: exactMatch.rates[region],
                    count: 1
                };
            }

            // Find the closest sizing by distance metric
            let bestMatch = null;
            let minDistance = Infinity;

            for (const inst of INSTANCE_CATALOG.rds) {
                // Since RAM is 4x CPU, normalize RAM by dividing by 4
                const distance = Math.abs(inst.cpu - cpuReq) + Math.abs(inst.ram - ramReq) / 4;
                if (distance < minDistance) {
                    minDistance = distance;
                    bestMatch = inst;
                } else if (distance === minDistance) {
                    // Tie-breaker: prefer the larger instance to avoid under-provisioning
                    if (inst.cpu >= cpuReq && (!bestMatch || bestMatch.cpu < cpuReq)) {
                        bestMatch = inst;
                    }
                }
            }

            return {
                name: bestMatch.name,
                rate: bestMatch.rates[region],
                count: 1
            };
        } else if (serviceType === 'msk') {
            // Kafka splits requirements into exactly 2 instances
            const halfCpu = cpuReq / 2;
            const halfRam = ramReq / 2;

            // Find an exact match first
            const exactMatch = INSTANCE_CATALOG.msk.find(inst => inst.cpu === halfCpu && inst.ram === halfRam);
            if (exactMatch) {
                return {
                    name: `2 x ${exactMatch.name}`,
                    rate: exactMatch.rates[region] * 2,
                    count: 1
                };
            }

            // Find the closest sizing by distance metric
            let bestMatch = null;
            let minDistance = Infinity;

            for (const inst of INSTANCE_CATALOG.msk) {
                const distance = Math.abs(inst.cpu - halfCpu) + Math.abs(inst.ram - halfRam) / 4;
                if (distance < minDistance) {
                    minDistance = distance;
                    bestMatch = inst;
                } else if (distance === minDistance) {
                    // Tie-breaker: prefer the larger instance to avoid under-provisioning
                    if (inst.cpu >= halfCpu && (!bestMatch || bestMatch.cpu < halfCpu)) {
                        bestMatch = inst;
                    }
                }
            }

            return {
                name: `2 x ${bestMatch.name}`,
                rate: bestMatch.rates[region] * 2,
                count: 1
            };
        } else if (serviceType === 'dms') {
            // Find an exact match first
            const exactMatch = INSTANCE_CATALOG.dms.find(inst => inst.cpu === cpuReq && inst.ram === ramReq);
            if (exactMatch) {
                return {
                    name: exactMatch.name,
                    rate: exactMatch.rates[region],
                    count: 1
                };
            }

            // Find the closest sizing by distance metric
            let bestMatch = null;
            let minDistance = Infinity;

            for (const inst of INSTANCE_CATALOG.dms) {
                const distance = Math.abs(inst.cpu - cpuReq) + Math.abs(inst.ram - ramReq) / 4;
                if (distance < minDistance) {
                    minDistance = distance;
                    bestMatch = inst;
                } else if (distance === minDistance) {
                    if (inst.cpu >= cpuReq && (!bestMatch || bestMatch.cpu < cpuReq)) {
                        bestMatch = inst;
                    }
                }
            }

            return {
                name: bestMatch.name,
                rate: bestMatch.rates[region],
                count: 1
            };
        } else {
            // Application and other: Split into smaller instances
            const selected = solveEC2Splitting(cpuReq, ramReq, region);
            const parts = [];
            let totalRate = 0;

            // Sort selected instances to display consistently (largest first)
            const sortedNames = Object.keys(selected).sort((a, b) => {
                const instA = INSTANCE_CATALOG.ec2.find(item => item.name === a);
                const instB = INSTANCE_CATALOG.ec2.find(item => item.name === b);
                return instB.cpu - instA.cpu || instB.ram - instA.ram;
            });

            sortedNames.forEach(name => {
                const count = selected[name];
                const inst = INSTANCE_CATALOG.ec2.find(item => item.name === name);
                parts.push(`${count} x ${name}`);
                totalRate += count * inst.rates[region];
            });

            return {
                name: parts.join(', '),
                rate: totalRate,
                count: 1
            };
        }
    };

    // Currency Formatting (VND and USD)
    const formatVND = (vnd) => {
        const val = (isNaN(vnd) || vnd === null || vnd === undefined) ? 0 : vnd;
        return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND', maximumFractionDigits: 0 }).format(val);
    };

    const formatUSD = (usd) => {
        const val = (isNaN(usd) || usd === null || usd === undefined) ? 0 : usd;
        return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(val);
    };

    // Currency conversion rate (1 USD = 26,000 VND)
    let USD_TO_VND = 26000;

    // 5. EXCEL EXPORTER / TEMPLATE GENERATOR
    const downloadTemplate = () => {
        const headers = [
            "STT", "Mã PM", "Tên PM", "Đơn vị đầu mối nghiệp vụ", "Phân nhóm", 
            "Chi phí hạ tầng được phân bổ 1 năm", "Tên cấu phần", 
            "CPU", "RAM (GB)", "Storage (GB)", "Data Transfer Out (GB)"
        ];

        // High fidelity sample data to help user understand the formatting
        const sampleRows = [
            [1, "PM-01", "Hệ thống Core Banking", "Khối Dịch vụ Tài chính", "N1 — Ứng dụng web, quản trị nội bộ", 650000000, "Web Server Cluster", 4, 8, 150, 100],
            [2, "PM-01", "Hệ thống Core Banking", "Khối Dịch vụ Tài chính", "N1 — Ứng dụng web, quản trị nội bộ", 650000000, "Database Postgres High-Availability", 16, 64, 1500, 500],
            [3, "PM-02", "Cổng thông tin Khách hàng (Portal)", "Phòng Quan hệ Công chúng", "N2 — Kênh khách hàng, tác nghiệp phụ thuộc Core", 120000000, "Nền tảng CMS Frontend", 2, 4, 50, 50],
            [4, "PM-02", "Cổng thông tin Khách hàng (Portal)", "Phòng Quan hệ Công chúng", "N2 — Kênh khách hàng, tác nghiệp phụ thuộc Core", 120000000, "Cơ sở dữ liệu người dùng SQL", 4, 16, 250, 150],
            [5, "PM-03", "Hệ thống Quản trị Nhân sự (HRM)", "Khối Tổ chức Nhân sự", "N6 — Phần mềm thương mại đóng gói", 90000000, "API Gateway & Application Server", 2, 4, 30, 20],
            [6, "PM-03", "Hệ thống Quản trị Nhân sự (HRM)", "Khối Tổ chức Nhân sự", "N6 — Phần mềm thương mại đóng gói", 90000000, "Database HRM Postgres", 2, 8, 80, 40],
            [7, "PM-04", "Ứng dụng Di động m-Banking", "Khối Ngân hàng Số", "N3 — Thanh toán, thẻ, chứng khoán, ký số", 180000000, "Microservices Application Engine", 12, 24, 200, 300],
            [8, "PM-05", "Nền tảng Báo cáo Phân tích BI & AI", "Phòng Quản lý Dữ liệu", "N4 — Dữ liệu, báo cáo, xử lý theo lô", 1200000000, "Spark/BI Data Processing Node", 32, 128, 500, 1000],
            [9, "PM-05", "Nền tảng Báo cáo Phân tích BI & AI", "Phòng Quản lý Dữ liệu", "N4 — Dữ liệu, báo cáo, xử lý theo lô", 1200000000, "Data Warehouse Postgres Analytics", 64, 256, 4000, 2000]
        ];

        // Build workbook
        const wb = XLSX.utils.book_new();
        const ws_data = [headers, ...sampleRows];
        const ws = XLSX.utils.aoa_to_sheet(ws_data);

        // Styling widths (optional but nice)
        ws['!cols'] = [
            { wch: 6 },  // STT
            { wch: 10 }, // Mã PM
            { wch: 35 }, // Tên PM
            { wch: 30 }, // Đơn vị nghiệp vụ
            { wch: 30 }, // Phân nhóm
            { wch: 32 }, // Chi phí On-Prem
            { wch: 35 }, // Tên cấu phần
            { wch: 8 },  // CPU
            { wch: 12 }, // RAM
            { wch: 14 }, // Storage
            { wch: 22 }  // Data Transfer Out (GB)
        ];

        XLSX.utils.book_append_sheet(wb, ws, "Danh sach hạ tầng");
        XLSX.writeFile(wb, "Bieu_mau_nhap_lieu_cloud_calc.xlsx");
    };

    // 6. PROCESS EXCEL FILE IMPORT
    const handleFile = (file) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Parse rows into JSON
                const rows = XLSX.utils.sheet_to_json(worksheet);
                
                if (rows.length === 0) {
                    alert("File Excel rỗng! Vui lòng kiểm tra lại dữ liệu.");
                    return;
                }

                // Verify headers by testing existence of key attributes
                const sampleRow = rows[0];
                const headerKeys = Object.keys(sampleRow);
                
                // Simple Vietnamese columns checker
                const matches = headerKeys.filter(k => 
                    k.includes("Mã") || k.includes("Tên") || k.includes("phần") || k.includes("CPU") || k.includes("RAM") || k.includes("Storage")
                );

                if (matches.length < 3) {
                    alert("Cảnh báo: Định dạng cột không khớp chính xác. Hãy chắc chắn rằng bạn dùng đúng file mẫu để kết quả tính chính xác nhất.");
                }

                parsePortfolioFromRows(rows);
                
                // Show file loaded state
                elFileName.innerText = file.name;
                elFileBadge.style.display = 'flex';
                elDropzone.style.display = 'none';

                // Recalculate
                recalculateAll();

            } catch (err) {
                console.error("Error reading excel file:", err);
                alert("Đã xảy ra lỗi khi đọc file Excel. Vui lòng thử lại.");
            }
        };
        reader.readAsArrayBuffer(file);
    };

    // Fuzzy matching helper to retrieve row values under slightly different spellings or capitalizations
    const getRowValue = (row, candidates) => {
        const keys = Object.keys(row);
        for (const key of keys) {
            // Normalize key: trim, lowercase, and strip Vietnamese accents
            const cleanKey = key.trim().toLowerCase()
                .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
                .replace(/đ/g, "d"); // replace special Vietnamese 'd'
                
            for (const candidate of candidates) {
                const cleanCandidate = candidate.trim().toLowerCase()
                    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
                    .replace(/đ/g, "d");
                    
                if (cleanKey === cleanCandidate || cleanKey.includes(cleanCandidate) || cleanCandidate.includes(cleanKey)) {
                    return row[key];
                }
            }
        }
        return "";
    };

    // Parse flat Excel rows into aggregated Applications
    const parsePortfolioFromRows = (rows) => {
        // Reset sorting states on new files
        currentSortColumn = 'stt';
        currentSortDirection = 'asc';

        const appMap = {};

        // Track last seen parent application info to support merged cells or blank header cells
        let lastSeenCode = "";
        let lastSeenName = "";
        let lastSeenOwner = "";
        let lastSeenOnPremStr = "0";
        let lastSeenGroup = "";

        rows.forEach((row) => {
            // Read values using fuzzy header matcher
            const rawStt = getRowValue(row, ["stt", "no", "no.", "thu tu"]);
            let rawCode = String(getRowValue(row, ["ma pm", "ma phan mem", "app code", "pm code", "ma_pm"]) || "").trim();
            let rawName = String(getRowValue(row, ["ten pm", "ten phan mem", "app name", "pm name", "ten_pm"]) || "").trim();
            let rawGroup = String(getRowValue(row, ["phan nhom", "nhom", "app group", "group"]) || "").trim();
            let rawOwner = String(getRowValue(row, ["don vi dau moi nghiep vu", "dau moi", "nghiep vu", "chu so huu", "business owner", "owner", "don vi"]) || "").trim();
            let rawOnPremStr = String(getRowValue(row, ["chi phi ha tang duoc phan bo 1 nam", "chi phi ha tang", "chi phi 1 nam", "on-prem cost", "onprem cost", "budget", "chi phi"]) || "");
            const rawCompName = String(getRowValue(row, ["ten cau phan", "cau phan", "component name", "component", "device", "ten thiet bi"]) || "").trim();
            let cpuVal = parseFloat(getRowValue(row, ["cpu", "vcpu", "cores", "core"]));
            let ramVal = parseFloat(getRowValue(row, ["ram (gb)", "ram", "memory", "dung luong ram"]));
            let storageVal = parseFloat(getRowValue(row, ["storage (gb)", "storage", "disk", "dung luong", "o cung", "hdd", "ssd"]));
            let dtoVal = parseFloat(getRowValue(row, ["data transfer out (gb)", "data transfer out", "data transfer", "transfer out", "outbound transfer", "dto"]));

            const rawCpu = isNaN(cpuVal) ? 0 : cpuVal;
            const rawRam = isNaN(ramVal) ? 0 : ramVal;
            const rawStorage = isNaN(storageVal) ? 0 : storageVal;
            const rawDto = isNaN(dtoVal) ? 0 : dtoVal;

            // Carry forward values if current row has component data but missing parent app headers
            if (!rawCode && !rawName && rawCompName) {
                rawCode = lastSeenCode;
                rawName = lastSeenName;
                rawOwner = lastSeenOwner;
                rawGroup = lastSeenGroup;
                if (!rawOnPremStr) rawOnPremStr = lastSeenOnPremStr;
            } else {
                // Keep record of last valid application header values
                if (rawCode) lastSeenCode = rawCode;
                if (rawName) lastSeenName = rawName;
                if (rawOwner) lastSeenOwner = rawOwner;
                if (rawGroup) lastSeenGroup = rawGroup;
                if (rawOnPremStr) lastSeenOnPremStr = rawOnPremStr;
            }

            if (!rawCode && !rawName) return; // Skip truly empty rows

            // Clean OnPremise Cost
            const cleanedOnPrem = parseFloat((rawOnPremStr || "0").replace(/[^0-9.-]/g, '')) || 0;

            const appKey = rawCode || rawName;

            if (!appMap[appKey]) {
                const normGroup = normalizeGroup(rawGroup);
                appMap[appKey] = {
                    maPM: rawCode || "PM-N/A",
                    tenPM: rawName || "Ứng dụng chưa đặt tên",
                    owner: rawOwner || "Chưa xác định",
                    onPremCost: cleanedOnPrem,
                    phanNhom: normGroup,
                    components: []
                };
            }

            // Keep components list
            appMap[appKey].components.push({
                stt: rawStt,
                name: rawCompName || "Cấu phần không tên",
                cpu: rawCpu,
                ram: rawRam,
                storage: rawStorage,
                dataTransferOut: rawDto,
                // Assign Database or Compute service
                serviceType: classifyService(rawCompName)
            });
        });

        parsedPortfolio = Object.values(appMap);
    };

    // Classify service dynamically: DB if name has database keywords, else EC2/EKS Compute
    const classifyService = (compName) => {
        const name = compName.toLowerCase();
        if (
            name.includes("dms") || 
            name.includes("migration") || 
            name.includes("replication") || 
            name.includes("replicate")
        ) {
            return "dms";
        }
        if (
            name.includes("kafka") || 
            name.includes("msk") || 
            name.includes("mq") || 
            name.includes("message queue")
        ) {
            return "msk";
        }
        if (
            name.includes("db") || 
            name.includes("database") || 
            name.includes("postgres") || 
            name.includes("sql") || 
            name.includes("oracle") || 
            name.includes("mysql") || 
            name.includes("rds") ||
            name.includes("mariadb")
        ) {
            return "rds";
        }
        return "ec2";
    };

    // 7. CORE CALCULATION ENGINE
    const recalculateAll = () => {
        if (parsedPortfolio.length === 0) return;

        const region = elAwsRegion.value;
        const plan = elAwsPlan.value;
        const ancillaryPct = parseFloat(elAncillaryPct.value);
        
        // Compliance backup checkbox status
        const elComplianceBackup = document.getElementById('compliance-realtime-backup');
        const isRealtimeBackup = elComplianceBackup ? elComplianceBackup.checked : false;

        const elMigrationCost = document.getElementById('compliance-migration-cost');
        const isMigrationCostActive = elMigrationCost ? elMigrationCost.checked : false;

        // Get Plan discount multiplier
        let discountMultiplier = 1.0;
        if (plan === 'ri1') discountMultiplier = 0.70; // 30% off
        if (plan === 'ri3') discountMultiplier = 0.50; // 50% off

        const dtoRate = INSTANCE_CATALOG.dataTransferOutRate !== undefined ? INSTANCE_CATALOG.dataTransferOutRate : 0.12;

        processedPortfolio = parsedPortfolio.map((app, index) => {
            let totalBaseAwsUSD = 0;
            let totalComputeUSD = 0;
            let totalDatabaseUSD = 0;
            let totalStorageUSD = 0;
            let totalDataTransferUSD = 0;
            let totalReplicationComputeVND = 0;
            let totalReplicationDtoVND = 0;
            let totalRealtimeReplicationVND = 0;
            let hasCompute = false;

            const calculatedComponents = app.components.map((comp) => {
                // Match the AWS Instance
                const matched = matchInstance(comp.serviceType, comp.cpu, comp.ram, region);
                
                // Compute rates
                const hourlyRate = matched.rate * matched.count * discountMultiplier;
                const yearlyComputeUSD = hourlyRate * 24 * 365; // Year rate
                
                // Storage cost
                const storageType = comp.serviceType === 'rds' ? 'rds' : (comp.serviceType === 'msk' ? 'msk' : (comp.serviceType === 'dms' ? 'dms' : 'ebs'));
                const storageRate = INSTANCE_CATALOG.storage[storageType][region];
                const yearlyStorageUSD = comp.storage * storageRate * 12;

                // Data Transfer Out cost
                const yearlyDataTransferUSD = (comp.dataTransferOut || 0) * dtoRate * 12;

                // Compliance replication calculation (only database i.e. rds)
                let realtimeReplicationCostVND = 0;
                let ec2ReplicationVND = 0;
                let dmsReplicationVND = 0;
                let dtoReplicationVND = 0;

                if (isRealtimeBackup && comp.serviceType === 'rds') {
                    // 1. EC2 equivalent
                    const ec2Matched = matchInstance('ec2', comp.cpu, comp.ram, region);
                    const ec2HourlyRate = ec2Matched.rate * ec2Matched.count * discountMultiplier;
                    const ec2YearlyUSD = ec2HourlyRate * 24 * 365;
                    ec2ReplicationVND = ec2YearlyUSD * USD_TO_VND;

                    // 2. DMS sizing DMS = 1/4 sizing of database
                    const dmsMatched = matchInstance('dms', comp.cpu / 4, comp.ram / 4, region);
                    const dmsHourlyRate = dmsMatched.rate * dmsMatched.count * discountMultiplier;
                    const dmsYearlyUSD = dmsHourlyRate * 24 * 365;
                    dmsReplicationVND = dmsYearlyUSD * USD_TO_VND;

                    // 3. DTO: Tỷ lệ thay đổi/ngày (%DB) * dung lượng database * 30 ngày
                    const dbChangeRatePct = (GROUP_INFO[app.phanNhom || 'N1'] || GROUP_INFO['N1']).dbChangeRate / 100;
                    const monthlyDmsDtoGB = dbChangeRatePct * comp.storage * 30;
                    const yearlyDmsDtoUSD = monthlyDmsDtoGB * dtoRate * 12;
                    dtoReplicationVND = yearlyDmsDtoUSD * USD_TO_VND;

                    realtimeReplicationCostVND = ec2ReplicationVND + dmsReplicationVND + dtoReplicationVND;

                    // Accumulate replication totals
                    totalReplicationComputeVND += ec2ReplicationVND + dmsReplicationVND;
                    totalReplicationDtoVND += dtoReplicationVND;
                    totalRealtimeReplicationVND += realtimeReplicationCostVND;
                }

                const componentYearlyTotalUSD = yearlyComputeUSD + yearlyStorageUSD + yearlyDataTransferUSD;
                const totalYearlyCostVND = (componentYearlyTotalUSD * USD_TO_VND) + realtimeReplicationCostVND;

                // Sum up aggregates
                totalBaseAwsUSD += componentYearlyTotalUSD;
                totalStorageUSD += yearlyStorageUSD;
                totalDataTransferUSD += yearlyDataTransferUSD;

                if (comp.serviceType === 'rds' || comp.serviceType === 'msk' || comp.serviceType === 'dms') {
                    totalDatabaseUSD += yearlyComputeUSD;
                } else {
                    totalComputeUSD += yearlyComputeUSD;
                    hasCompute = true; // Mark that app uses computation
                }

                return {
                    ...comp,
                    matchedInstance: matched.name,
                    matchedInstanceCount: matched.count,
                    yearlyComputeCostVND: yearlyComputeUSD * USD_TO_VND,
                    yearlyStorageCostVND: yearlyStorageUSD * USD_TO_VND,
                    yearlyDataTransferCostVND: yearlyDataTransferUSD * USD_TO_VND,
                    realtimeReplicationCostVND,
                    ec2ReplicationVND,
                    dmsReplicationVND,
                    dtoReplicationVND,
                    totalYearlyCostVND
                };
            });

            // Add Cluster EKS Fee if EKS/EC2 is present
            let eksClusterFeeUSD = 0;
            let eksClusterFeeVND = 0;
            if (hasCompute) {
                eksClusterFeeUSD = INSTANCE_CATALOG.eksClusterYear;
                eksClusterFeeVND = eksClusterFeeUSD * USD_TO_VND;
                totalComputeUSD += eksClusterFeeUSD;
                totalBaseAwsUSD += eksClusterFeeUSD;
            }

            // Ancillary costs
            const totalAncillaryUSD = totalBaseAwsUSD * (ancillaryPct / 100);
            const totalAncillaryVND = totalAncillaryUSD * USD_TO_VND;

            // Migration complexity assessment and mandays lookup
            const numComponents = app.components.length;
            let matchedComplexity = COMPLEXITY_INFO.find(item => numComponents >= item.minComp && numComponents <= item.maxComp);
            if (!matchedComplexity && numComponents > 0) {
                matchedComplexity = COMPLEXITY_INFO[COMPLEXITY_INFO.length - 1];
            }
            const migrationMandays = numComponents > 0 ? (matchedComplexity ? matchedComplexity.mandays : 0) : 0;
            const complexityLevel = numComponents > 0 ? (matchedComplexity ? matchedComplexity.level : "Không xác định") : "Không có";
            const migrationCostVND = isMigrationCostActive ? (migrationMandays * MANDAY_RATE) : 0;

            // Grand total
            const grandTotalAwsUSD = totalBaseAwsUSD + totalAncillaryUSD;
            const grandTotalAwsVND = (grandTotalAwsUSD * USD_TO_VND) + totalRealtimeReplicationVND + migrationCostVND;

            // Deltas
            const deltaVND = grandTotalAwsVND - app.onPremCost;
            const deltaPct = app.onPremCost > 0 ? (deltaVND / app.onPremCost) * 100 : 0;

            // Get stability & performance properties from Group dictionary
            const grp = GROUP_INFO[app.phanNhom || "N1"] || GROUP_INFO["N1"];
            const stability = grp.stability;
            const performance = grp.performance;
            const stabilityAbbr = grp.stabilityAbbr || "M1";
            const performanceAbbr = grp.performanceAbbr || "M1";

            // Default algorithmic recommendation: Nên if within 15% delta threshold
            const defaultRecommend = grandTotalAwsVND <= (app.onPremCost * 1.15);

            return {
                ...app,
                originalIndex: index,
                components: calculatedComponents,
                eksClusterFeeVND,
                totalAncillaryVND,
                totalDataTransferVND: (totalDataTransferUSD * USD_TO_VND) + totalReplicationDtoVND,
                grandTotalAwsVND,
                totalComputeVND: totalComputeUSD * USD_TO_VND,
                totalDatabaseVND: (totalDatabaseUSD * USD_TO_VND) + totalReplicationComputeVND,
                totalStorageVND: totalStorageUSD * USD_TO_VND,
                totalReplicationComputeVND,
                totalReplicationDtoVND,
                totalRealtimeReplicationVND,
                migrationMandays,
                migrationCostVND,
                complexityLevel,
                deltaVND,
                deltaPct,
                stability,
                performance,
                stabilityAbbr,
                performanceAbbr,
                // Recommendations (allow local overrides)
                recommend: defaultRecommend,
                isOverridden: false
            };
        });

        // Apply active sorting
        sortPortfolio();

        // Update Stats UI & Table view
        renderDashboard();
    };

    // 8. RENDER DASHBOARD (Stats, Charts, and Table)
    const renderDashboard = () => {
        if (processedPortfolio.length === 0) return;

        // Enable Export/Print
        elBtnExportExcel.removeAttribute('disabled');
        elBtnPrintReport.removeAttribute('disabled');

        // A. Calc Top Stats
        let totalOnPrem = 0;
        let totalAws = 0;
        let totalComponents = 0;

        processedPortfolio.forEach(app => {
            totalOnPrem += app.onPremCost;
            totalAws += app.grandTotalAwsVND;
            totalComponents += app.components.length;
        });

        elStatTotalApps.innerText = processedPortfolio.length;
        elStatTotalComponents.innerText = totalComponents;
        elStatOnPremCost.innerText = formatVND(totalOnPrem);
        elStatAwsCost.innerText = formatVND(totalAws);

        // Update Savings Icon
        const savingsDelta = totalOnPrem - totalAws;
        if (savingsDelta >= 0) {
            elStatAwsCostIcon.className = "stat-icon-wrapper neon-red";
            elStatAwsCostIcon.innerHTML = `<i data-lucide="trending-up"></i>`;
            elStatAwsCost.className = "stat-value text-glow-red";
        }

        lucide.createIcons();

        // B. Render Table Rows
        renderTableRows();

        // C. Render Charts
        updateChartsData();
    };

    // Render primary comparison grid records
    const renderTableRows = () => {
        const searchTerm = elSearchInput.value.toLowerCase().trim();
        elTableBody.innerHTML = '';

        const filtered = processedPortfolio.filter(app => 
            app.maPM.toLowerCase().includes(searchTerm) ||
            app.tenPM.toLowerCase().includes(searchTerm) ||
            app.owner.toLowerCase().includes(searchTerm)
        );

        if (filtered.length === 0) {
            elTableBody.innerHTML = `
                <tr>
                    <td colspan="14" class="empty-table-state">
                        <div class="empty-state-content">
                            <i data-lucide="search"></i>
                            <h4>Không tìm thấy kết quả</h4>
                            <p>Thử điều chỉnh từ khóa tìm kiếm hoặc kiểm tra lại file dữ liệu.</p>
                        </div>
                    </td>
                </tr>
            `;
            lucide.createIcons();
            return;
        }

        filtered.forEach((app, index) => {
            // Breakdowns metrics
            const computePct = Math.round((app.totalComputeVND / app.grandTotalAwsVND) * 100) || 0;
            const dbPct = Math.round((app.totalDatabaseVND / app.grandTotalAwsVND) * 100) || 0;
            const storagePct = Math.round((app.totalStorageVND / app.grandTotalAwsVND) * 100) || 0;
            const totalPhuTroVND = app.totalAncillaryVND + (app.totalDataTransferVND || 0);
            const ancillaryPct = Math.round((totalPhuTroVND / app.grandTotalAwsVND) * 100) || 0;

            const isSave = app.deltaVND <= 0;
            const deltaClass = isSave ? 'delta-save' : 'delta-increase';
            const deltaSymbol = isSave ? '-' : '+';
            const deltaText = `${deltaSymbol}${formatVND(Math.abs(app.deltaVND))}`;
            const pctText = `${deltaSymbol}${Math.abs(app.deltaPct).toFixed(1)}%`;

            // Recommendation Badge
            const recClass = app.recommend ? 'rec-should' : 'rec-should-not';
            const recIcon = app.recommend ? 'check-circle' : 'alert-triangle';
            const recLabel = app.recommend ? 'Có' : 'Không';

            // HTML Master Row
            const trMain = document.createElement('tr');
            trMain.className = 'table-main-row';
            trMain.setAttribute('data-app-id', app.maPM);
            trMain.innerHTML = `
                <td>${index + 1}</td>
                <td><span class="pm-code">${app.maPM}</span></td>
                <td><strong>${app.tenPM}</strong></td>
                <td><span class="biz-owner">${app.owner}</span></td>
                <td><span class="group-badge group-${app.phanNhom}" title="${GROUP_INFO[app.phanNhom].name}">${app.phanNhom}</span></td>
                <td class="text-right cost-bold-yellow">${formatVND(app.onPremCost)}</td>
                <td class="text-right cost-bold-green">${formatVND(app.grandTotalAwsVND)}</td>
                <td>
                    <div class="service-distribution-bar">
                        <div class="dist-bar-track" title="Compute: ${computePct}%, DB: ${dbPct}%, Storage: ${storagePct}%, Khác: ${ancillaryPct}%">
                            <div class="dist-bar-item dist-bar-compute" style="width: ${computePct}%"></div>
                            <div class="dist-bar-item dist-bar-db" style="width: ${dbPct}%"></div>
                            <div class="dist-bar-item dist-bar-other" style="width: ${storagePct + ancillaryPct}%"></div>
                        </div>
                        <div class="dist-labels">
                            <span class="dist-label-item"><span class="dist-dot dist-bar-compute"></span> Comp: ${computePct}%</span>
                            <span class="dist-label-item"><span class="dist-dot dist-bar-db"></span> DB: ${dbPct}%</span>
                            <span class="dist-label-item"><span class="dist-dot dist-bar-other"></span> Khác: ${storagePct + ancillaryPct}%</span>
                        </div>
                    </div>
                </td>
                <td class="text-right">
                    <span class="delta-amount ${deltaClass}">${deltaText}</span>
                </td>
                <td class="text-right">
                    <span class="delta-pct ${deltaClass}" style="font-weight: 700; font-size: 0.85rem; font-family: var(--font-display);">${pctText}</span>
                </td>
                <td style="text-align: center;"><span class="text-subtle-desc cursor-help" title="${app.stability}">${app.stabilityAbbr}</span></td>
                <td style="text-align: center;"><span class="text-subtle-desc cursor-help" title="${app.performance}">${app.performanceAbbr}</span></td>
                <td style="text-align: center;">
                    <span class="recommendation-badge ${recClass}">
                        ${recLabel}
                    </span>
                </td>
                <td class="text-right">
                    <button class="btn-arrow btn-icon" title="Xem chi tiết"><i data-lucide="chevron-down"></i></button>
                </td>
            `;

            // HTML Expandable detail row drawer
            const trDetail = document.createElement('tr');
            trDetail.className = 'detail-row';
            trDetail.style.display = 'none'; // Initially hidden
            trDetail.innerHTML = `
                <td colspan="14">
                    <div class="detail-drawer-container">
                        <div class="drawer-grid">
                            <div class="components-table-block">
                                <div class="drawer-title-row">
                                    <h4><i data-lucide="cpu"></i> Cấu phần tài nguyên chi tiết & Ánh xạ AWS</h4>
                                    <span>${app.components.length} Cấu phần được nhập</span>
                                </div>
                                <table class="drawer-table">
                                    <thead>
                                        <tr>
                                            <th>Cấu phần</th>
                                            <th>Loại</th>
                                            <th class="text-right">vCPU</th>
                                            <th class="text-right">RAM (GB)</th>
                                            <th class="text-right">Storage (GB)</th>
                                            <th class="text-right">Data Transfer (GB)</th>
                                            <th>AWS Dịch vụ quy đổi</th>
                                            <th class="text-right">Phí Compute/Năm</th>
                                            <th class="text-right">Phí Storage/Năm</th>
                                            <th class="text-right">Phí Data Transfer/Năm</th>
                                            <th class="text-right">Tổng chi phí/Năm</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        ${app.components.map(comp => `
                                            <tr>
                                                <td><strong>${comp.name}</strong></td>
                                                <td><span class="service-pill ${comp.serviceType === 'rds' ? 'db-pill' : (comp.serviceType === 'msk' ? 'msk-pill' : (comp.serviceType === 'dms' ? 'dms-pill' : 'app-pill'))}">${comp.serviceType === 'rds' ? 'RDS Postgres' : (comp.serviceType === 'msk' ? 'MSK Kafka' : (comp.serviceType === 'dms' ? 'DMS Migration' : 'EKS/EC2'))}</span></td>
                                                <td class="text-right">${comp.cpu}</td>
                                                <td class="text-right">${comp.ram} GB</td>
                                                <td class="text-right">${comp.storage} GB</td>
                                                <td class="text-right">${comp.dataTransferOut || 0} GB</td>
                                                <td><span class="instance-code">${comp.matchedInstanceCount === 1 ? comp.matchedInstance : comp.matchedInstanceCount + ' x ' + comp.matchedInstance}</span></td>
                                                <td class="text-right cost-bold-blue">${formatVND(comp.yearlyComputeCostVND)}</td>
                                                <td class="text-right cost-bold-yellow">${formatVND(comp.yearlyStorageCostVND)}</td>
                                                <td class="text-right cost-bold-purple" style="color: #8E24AA !important;">${formatVND(comp.yearlyDataTransferCostVND || 0)}</td>
                                                <td class="text-right cost-bold-green">${formatVND(comp.totalYearlyCostVND - (comp.realtimeReplicationCostVND || 0))}</td>
                                            </tr>
                                            ${comp.realtimeReplicationCostVND > 0 ? `
                                                <tr style="background: rgba(0, 90, 54, 0.015);">
                                                    <td style="padding-left: 20px;"><i data-lucide="corner-down-right" style="width: 12px; height: 12px; display: inline-block; vertical-align: middle; margin-right: 5px; color: var(--primary-color);"></i> <span style="font-size: 0.76rem; color: var(--text-muted);">Bản sao cập nhật real-time VN cho <strong>${comp.name}</strong></span></td>
                                                    <td><span class="service-pill" style="background: rgba(0, 90, 54, 0.05); color: var(--primary-color); border: 1px solid rgba(0, 90, 54, 0.1);">Replication (VN)</span></td>
                                                    <td class="text-right">-</td>
                                                    <td class="text-right">-</td>
                                                    <td class="text-right">-</td>
                                                    <td class="text-right">-</td>
                                                    <td><span class="instance-code" title="EC2 Sizing tương đương + DMS 1/4 Sizing + DTO %DB">EC2 + DMS + DTO</span></td>
                                                    <td class="text-right cost-bold-blue" title="EC2: ${formatVND(comp.ec2ReplicationVND)} + DMS: ${formatVND(comp.dmsReplicationVND)}">${formatVND(comp.ec2ReplicationVND + comp.dmsReplicationVND)}</td>
                                                    <td class="text-right">-</td>
                                                    <td class="text-right cost-bold-purple" title="DTO %DB: ${formatVND(comp.dtoReplicationVND)}" style="color: #8E24AA !important;">${formatVND(comp.dtoReplicationVND)}</td>
                                                    <td class="text-right cost-bold-green" style="font-weight: 600;">${formatVND(comp.realtimeReplicationCostVND)}</td>
                                                </tr>
                                            ` : ''}
                                        `).join('')}
                                        ${app.eksClusterFeeVND > 0 ? `
                                            <tr>
                                                <td><strong>Phí quản lý EKS Cluster</strong></td>
                                                <td><span class="service-pill app-pill">EKS Cluster</span></td>
                                                <td class="text-right">-</td>
                                                <td class="text-right">-</td>
                                                <td class="text-right">-</td>
                                                <td class="text-right">-</td>
                                                <td><span class="instance-code">eks.controlplane</span></td>
                                                <td class="text-right cost-bold-blue">${formatVND(app.eksClusterFeeVND)}</td>
                                                <td class="text-right">-</td>
                                                <td class="text-right">-</td>
                                                <td class="text-right cost-bold-green">${formatVND(app.eksClusterFeeVND)}</td>
                                            </tr>
                                        ` : ''}
                                        <tr>
                                            <td><strong>Chi phí dịch vụ bổ trợ</strong></td>
                                            <td><span class="service-pill" style="background: rgba(0, 90, 54, 0.08); color: #005A36; border: 1px solid rgba(0, 90, 54, 0.15);">Ancillary</span></td>
                                            <td class="text-right">-</td>
                                            <td class="text-right">-</td>
                                            <td class="text-right">-</td>
                                            <td class="text-right">-</td>
                                            <td><span class="instance-code">Network & Ops (${elAncillaryPct.value}%)</span></td>
                                            <td class="text-right">-</td>
                                            <td class="text-right">-</td>
                                            <td class="text-right">-</td>
                                            <td class="text-right cost-bold-green">${formatVND(app.totalAncillaryVND)}</td>
                                        </tr>
                                        ${app.totalRealtimeReplicationVND > 0 ? `
                                            <tr style="background: rgba(0, 90, 54, 0.02); font-weight:600;">
                                                <td colspan="6">Trong đó: Phí tuân thủ sao lưu Việt Nam (Real-time)</td>
                                                <td><span class="instance-code" style="color: var(--primary-color);">Compliance</span></td>
                                                <td class="text-right" style="color: #1565C0;">${formatVND(app.totalReplicationComputeVND)}</td>
                                                <td class="text-right">-</td>
                                                <td class="text-right" style="color: #8E24AA;">${formatVND(app.totalReplicationDtoVND)}</td>
                                                <td class="text-right" style="color: var(--primary-color);">${formatVND(app.totalRealtimeReplicationVND)}</td>
                                            </tr>
                                        ` : ''}
                                         ${app.migrationCostVND > 0 ? `
                                             <tr style="background: rgba(0, 90, 54, 0.02); font-weight:600;">
                                                 <td colspan="6">Chi phí chuyển đổi (Migration Mandays: <strong>${app.migrationMandays} ngày công</strong>)</td>
                                                 <td><span class="instance-code" style="color: var(--primary-color);">${app.complexityLevel}</span></td>
                                                 <td class="text-right">-</td>
                                                 <td class="text-right">-</td>
                                                 <td class="text-right">-</td>
                                                 <td class="text-right" style="color: var(--primary-color); font-weight: 700;">${formatVND(app.migrationCostVND)}</td>
                                             </tr>
                                         ` : ''}
                                         <tr style="background: rgba(0, 90, 54, 0.04); font-weight:700;">
                                            <td colspan="6">Tổng cộng toàn bộ cấu phần ứng dụng</td>
                                            <td><span class="instance-code" style="color: #005A36;">AWS Mapped</span></td>
                                            <td class="text-right" style="color: #1565C0;">${formatVND(app.totalComputeVND + app.eksClusterFeeVND + app.totalReplicationComputeVND)}</td>
                                            <td class="text-right" style="color: #FDB813;">${formatVND(app.totalStorageVND)}</td>
                                            <td class="text-right" style="color: #8E24AA;">${formatVND(app.totalDataTransferVND || 0)}</td>
                                            <td class="text-right" style="color: #005A36;">${formatVND(app.grandTotalAwsVND)}</td>
                                        </tr>
                                    </tbody>
                                </table>
                            </div>
                            <div class="decision-block">
                                <div class="drawer-title-row">
                                    <h4><i data-lucide="brain"></i> Đánh giá & Khuyến nghị tối ưu</h4>
                                    <span>Mã PM: ${app.maPM}</span>
                                </div>
                                <div class="drawer-rec-card">
                                    <div class="drawer-rec-title">
                                        <i data-lucide="${recIcon}" style="color: ${app.recommend ? 'var(--neon-green)' : 'var(--neon-red)'}"></i>
                                        <span>Phân tích hiệu quả kinh tế</span>
                                    </div>
                                    <p class="drawer-rec-desc">
                                        ${getDetailedRecommendationText(app)}
                                    </p>
                                    <div class="drawer-rec-action-row">
                                        <span class="rec-indicator-text ${app.recommend ? 'should' : 'shouldnot'}">
                                            Khuyến nghị: ${app.recommend ? 'NÊN DI TRÚ' : 'KHÔNG NÊN'}
                                        </span>
                                        <label class="switch" title="Ghi đè đề xuất hệ thống">
                                            <input type="checkbox" class="rec-override-toggle" ${app.recommend ? 'checked' : ''}>
                                            <span class="slider-toggle"></span>
                                        </label>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </td>
            `;

            elTableBody.appendChild(trMain);
            elTableBody.appendChild(trDetail);

            trMain.addEventListener('click', (e) => {
                if (e.target.closest('.switch') || e.target.closest('.rec-override-toggle')) return;
                const isOpen = trDetail.style.display !== 'none';
                document.querySelectorAll('.detail-row').forEach(row => row.style.display = 'none');
                document.querySelectorAll('.table-main-row').forEach(row => row.classList.remove('active-expanded'));
                if (!isOpen) {
                    trDetail.style.display = 'table-row';
                    const container = trDetail.querySelector('.detail-drawer-container');
                    container.style.display = 'block';
                    trMain.classList.add('active-expanded');
                }
            });

            const overrideToggle = trDetail.querySelector('.rec-override-toggle');
            overrideToggle.addEventListener('change', (e) => {
                const checked = e.target.checked;
                app.recommend = checked;
                const badge = trMain.querySelector('.recommendation-badge');
                badge.className = `recommendation-badge ${checked ? 'rec-should' : 'rec-should-not'}`;
                badge.innerHTML = `${checked ? 'Nên đưa lên Cloud' : 'Không nên đưa lên Cloud'}`;
                const indicatorText = trDetail.querySelector('.rec-indicator-text');
                indicatorText.className = `rec-indicator-text ${checked ? 'should' : 'shouldnot'}`;
                indicatorText.innerText = `Khuyến nghị: ${checked ? 'NÊN DI TRÚ' : 'KHÔNG NÊN'}`;
                const titleIcon = trDetail.querySelector('.drawer-rec-title i');
                titleIcon.style.color = checked ? 'var(--neon-green)' : 'var(--neon-red)';
                titleIcon.setAttribute('data-lucide', checked ? 'check-circle' : 'alert-triangle');
                lucide.createIcons();
            });
        });
        lucide.createIcons();
    };

    const getDetailedRecommendationText = (app) => {
        const diff = app.onPremCost - app.grandTotalAwsVND;
        const absDiffText = formatVND(Math.abs(diff));
        const absPctText = Math.abs(app.deltaPct).toFixed(1) + "%";
        if (diff >= 0) {
            return `Chi phi vận hành hệ thống <strong>${app.tenPM}</strong> trên đám mây AWS ước tính giúp doanh nghiệp tiết kiệm <strong>${absDiffText}</strong> mỗi năm (giảm <strong>${absPctText}</strong> so với On-Premise). 
            Việc dịch chuyển là cực kỳ tối ưu về mặt kinh tế, đồng thời giúp loại bỏ gánh nặng vận hành phần cứng vật lý tại chỗ. 
            <br><br>
            <strong>Đề xuất:</strong> Thực hiện kế hoạch nâng dịch vụ (Rehost/Replatform) lên cụm máy chủ container hoặc cơ sở dữ liệu Postgres tự động hóa của AWS.`;
        } else {
            return `Chi phi chạy ứng dụng <strong>${app.tenPM}</strong> trên AWS ước tính đang cao hơn ngân sách On-Premise cũ khoảng <strong>${absDiffText}</strong> (tăng <strong>${absPctText}</strong>). 
            Điều này xảy ra do lượng phần cứng yêu cầu tương đối lớn trong khi hạ tầng On-Premise hiện tại được phân bổ ngân sách khá hạn chế. 
            <br><br>
            <strong>Đề xuất tối ưu hóa:</strong> Doanh nghiệp có thể áp dụng <strong>Reserved Instances (1 năm / 3 năm)</strong> để giảm ngay tới 30-50% chi phí máy chủ, hoặc thực hiện tối ưu tài nguyên (Right-sizing) trước khi đưa lên Cloud để thu hẹp khoảng cách chi phí.`;
        }
    };

    // 9. RE-RENDER ANALYTICAL DASHBOARD CHARTS
    const updateChartsData = () => {
        const sortedApps = [...processedPortfolio].sort((a, b) => b.grandTotalAwsVND - a.grandTotalAwsVND).slice(0, 10);
        const labels = sortedApps.map(app => app.maPM);
        const onPremCosts = sortedApps.map(app => app.onPremCost);
        const awsCosts = sortedApps.map(app => app.grandTotalAwsVND);

        if (onPremVsAwsChart) onPremVsAwsChart.destroy();

        const ctxBar = document.getElementById('chart-comparison').getContext('2d');
        onPremVsAwsChart = new Chart(ctxBar, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [
                    {
                        label: 'Chi phí On-Premise (Năm)',
                        data: onPremCosts,
                        backgroundColor: 'rgba(120, 144, 156, 0.65)',
                        borderColor: '#78909C',
                        borderWidth: 2,
                        borderRadius: 6
                    },
                    {
                        label: 'Chi phí AWS Cloud (Năm)',
                        data: awsCosts,
                        backgroundColor: 'rgba(0, 90, 54, 0.75)',
                        borderColor: '#005A36',
                        borderWidth: 2,
                        borderRadius: 6
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'top',
                        labels: { color: '#0F291E', font: { family: 'Inter', size: 11, weight: '600' } }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                return ` ${context.dataset.label}: ${formatVND(context.raw)}`;
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        grid: { color: 'rgba(15, 41, 30, 0.08)' },
                        ticks: { color: '#5A6E65', font: { family: 'Inter' } }
                    },
                    x: {
                        grid: { display: false },
                        ticks: { color: '#5A6E65', font: { family: 'Inter' } }
                    }
                }
            }
        });

        let totalCompute = 0;
        let totalDatabase = 0;
        let totalStorage = 0;
        let totalAncillary = 0;

        processedPortfolio.forEach(app => {
            totalCompute += (app.totalComputeVND + app.eksClusterFeeVND);
            totalDatabase += app.totalDatabaseVND;
            totalStorage += app.totalStorageVND;
            totalAncillary += app.totalAncillaryVND;
        });

        if (awsBreakdownChart) awsBreakdownChart.destroy();

        const ctxPie = document.getElementById('chart-distribution').getContext('2d');
        awsBreakdownChart = new Chart(ctxPie, {
            type: 'doughnut',
            data: {
                labels: ['EKS/EC2 Compute', 'RDS Postgres DB', 'Lưu trữ SSD (gp3)', 'Dịch vụ phụ trợ'],
                datasets: [{
                    data: [totalCompute, totalDatabase, totalStorage, totalAncillary],
                    backgroundColor: [
                        '#005A36',  // EKS/EC2 Compute (BIDV Emerald Green)
                        '#1565C0',  // RDS Postgres DB (Corporate Royal Blue)
                        '#FDB813',  // Lưu trữ SSD (gp3) (BIDV Golden Apricot)
                        '#8E24AA'   // Dịch vụ phụ trợ (Elegant Amethyst Purple)
                    ],
                    borderWidth: 0,
                    hoverOffset: 10
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'right',
                        labels: { color: '#0F291E', font: { family: 'Inter', size: 11, weight: '600' } }
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                const percentage = total > 0 ? ((context.raw / total) * 100).toFixed(1) : "0.0";
                                return ` ${context.label}: ${formatVND(context.raw)} (${percentage}%)`;
                            }
                        }
                    }
                },
                cutout: '65%'
            }
        });
    };

    // 10. DYNAMIC SPREADSHEET EXPORTER FOR CALCULATED RESULTS
    const exportResultsToExcel = () => {
        if (processedPortfolio.length === 0) return;

        // 1. Generate Summary Data for Sheet 1
        const summaryData = processedPortfolio.map((app, index) => {
            const computePct = Math.round((app.totalComputeVND / app.grandTotalAwsVND) * 100) || 0;
            const dbPct = Math.round((app.totalDatabaseVND / app.grandTotalAwsVND) * 100) || 0;
            const storagePct = Math.round((app.totalStorageVND / app.grandTotalAwsVND) * 100) || 0;
            const totalPhuTroVND = app.totalAncillaryVND + (app.totalDataTransferVND || 0);
            const ancillaryPct = Math.round((totalPhuTroVND / app.grandTotalAwsVND) * 100) || 0;

            const recText = app.recommend ? "Có" : "Không";
            const deltaAmtText = `${app.deltaVND <= 0 ? 'Giảm' : 'Tăng'} ${Math.abs(app.deltaVND).toLocaleString()}đ`;
            const deltaPctText = `${app.deltaVND <= 0 ? '-' : '+'}${Math.abs(app.deltaPct).toFixed(2)}%`;

            const serviceBreakdown = `Compute: ${app.totalComputeVND.toLocaleString()}đ (${computePct}%), DB: ${app.totalDatabaseVND.toLocaleString()}đ (${dbPct}%), Storage: ${app.totalStorageVND.toLocaleString()}đ (${storagePct}%), Phụ trợ: ${totalPhuTroVND.toLocaleString()}đ (${ancillaryPct}%)`;

            return {
                "STT": index + 1,
                "Mã PM": app.maPM,
                "Tên PM": app.tenPM,
                "Đơn vị đầu mối nghiệp vụ": app.owner,
                "Phân nhóm": GROUP_INFO[app.phanNhom].name,
                "Chi phí hạ tầng được phân bỏ 1 năm": Math.round(app.onPremCost),
                "Chi phí trên AWS Cloud trong 1 năm": Math.round(app.grandTotalAwsVND),
                "Phí tuân thủ sao lưu Việt Nam (VND)": Math.round(app.totalRealtimeReplicationVND || 0),
                "Mức độ phức tạp": app.complexityLevel,
                "Số ngày công chuyển đổi (Mandays)": app.migrationMandays,
                "Chi phí chuyển đổi (VND)": Math.round(app.migrationCostVND || 0),
                "Số tiền/Tỷ lệ % dịch vụ": serviceBreakdown,
                "Số tiền tăng/giảm": deltaAmtText,
                "Tỷ lệ tăng/giảm": deltaPctText,
                "Độ ổn định": app.stability,
                "Hiệu năng": app.performance,
                "Đề xuất đưa lên Cloud": recText
            };
        });

        // 2. Generate Detailed Component-level Data for Sheet 2
        const exportData = [];
        let globalCompStt = 1;

        processedPortfolio.forEach(app => {
            app.components.forEach(comp => {
                exportData.push({
                    "STT": globalCompStt++,
                    "Mã PM": app.maPM,
                    "Tên PM": app.tenPM,
                    "Đơn vị đầu mối nghiệp vụ": app.owner,
                    "Phân nhóm": GROUP_INFO[app.phanNhom].name,
                    "Độ ổn định": app.stability,
                    "Hiệu năng": app.performance,
                    "Tên cấu phần": comp.name,
                    "Loại cấu phần": comp.serviceType === 'rds' ? "RDS Postgres" : (comp.serviceType === 'msk' ? "MSK Kafka" : (comp.serviceType === 'dms' ? "DMS Migration" : "EKS/EC2")),
                    "CPU": comp.cpu,
                    "RAM (GB)": comp.ram,
                    "Storage (GB)": comp.storage,
                    "Data Transfer Out (GB)": comp.dataTransferOut || 0,
                    "AWS Dịch vụ quy đổi": comp.matchedInstanceCount === 1 ? comp.matchedInstance : `${comp.matchedInstanceCount} x ${comp.matchedInstance}`,
                    "Phí AWS Compute / Năm (VND)": Math.round(comp.yearlyComputeCostVND),
                    "Phí AWS Storage / Năm (VND)": Math.round(comp.yearlyStorageCostVND),
                    "Phí AWS Data Transfer / Năm (VND)": Math.round(comp.yearlyDataTransferCostVND || 0),
                    "Phí tuân thủ sao lưu Việt Nam (VND)": Math.round(comp.realtimeReplicationCostVND || 0),
                    "Tổng chi phí AWS Cấu phần / Năm (VND)": Math.round(comp.totalYearlyCostVND)
                });
            });
        });

        // Generate workbook
        const wb = XLSX.utils.book_new();

        // Append Sheet 1
        const wsSummary = XLSX.utils.json_to_sheet(summaryData);
        wsSummary['!cols'] = [
            { wch: 6 },   // STT
            { wch: 12 },  // Mã PM
            { wch: 30 },  // Tên PM
            { wch: 25 },  // Đơn vị đầu mối nghiệp vụ
            { wch: 35 },  // Phân nhóm
            { wch: 35 },  // Chi phí hạ tầng được phân bổ 1 năm
            { wch: 35 },  // Chi phí trên AWS Cloud trong 1 năm
            { wch: 35 },  // Phí tuân thủ sao lưu Việt Nam (VND)
            { wch: 25 },  // Mức độ phức tạp
            { wch: 30 },  // Số ngày công chuyển đổi (Mandays)
            { wch: 25 },  // Chi phí chuyển đổi (VND)
            { wch: 80 },  // Số tiền/Tỷ lệ % dịch vụ
            { wch: 25 },  // Số tiền tăng/giảm
            { wch: 20 },  // Tỷ lệ tăng/giảm
            { wch: 45 },  // Độ ổn định
            { wch: 45 },  // Hiệu năng
            { wch: 30 }   // Đề xuất
        ];
        XLSX.utils.book_append_sheet(wb, wsSummary, "Tong_hop_PM");

        // Append Sheet 2
        const wsDetail = XLSX.utils.json_to_sheet(exportData);
        wsDetail['!cols'] = [
            { wch: 6 },   // STT
            { wch: 12 },  // Mã PM
            { wch: 30 },  // Tên PM
            { wch: 25 },  // Đơn vị đầu mối nghiệp vụ
            { wch: 35 },  // Phân nhóm
            { wch: 45 },  // Độ ổn định
            { wch: 45 },  // Hiệu năng
            { wch: 30 },  // Tên cấu phần
            { wch: 15 },  // Loại cấu phần
            { wch: 8 },   // CPU
            { wch: 12 },  // RAM
            { wch: 14 },  // Storage
            { wch: 22 },  // Data Transfer Out (GB)
            { wch: 25 },  // AWS Dịch vụ quy đổi
            { wch: 28 },  // Phí AWS Compute
            { wch: 28 },  // Phí AWS Storage
            { wch: 32 },  // Phí AWS Data Transfer
            { wch: 35 },  // Phí tuân thủ sao lưu Việt Nam (VND)
            { wch: 32 }   // Tổng chi phí
        ];
        XLSX.utils.book_append_sheet(wb, wsDetail, "Chi_tiet_cau_phan");

        XLSX.writeFile(wb, "Bao_cao_chi_phi_AWS_Cloud_Migration.xlsx");
    };

    // 10b. SYNC LIVE PRICES FROM AWS API WITH CORS PROXY
    const syncLivePricing = async () => {
        const btnSync = document.getElementById('btn-sync-pricing');
        if (!btnSync) return;

        const originalHTML = btnSync.innerHTML;
        btnSync.setAttribute('disabled', 'true');
        btnSync.innerHTML = `<span class="spinner"></span> Đang đồng bộ...`;

        const region = elAwsRegion.value;
        const proxyUrl = "https://api.allorigins.win/raw?url=";
        const ec2Url = `${proxyUrl}${encodeURIComponent(`https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonEC2/current/${region}/index.json`)}`;
        const rdsUrl = `${proxyUrl}${encodeURIComponent(`https://pricing.us-east-1.amazonaws.com/offers/v1.0/aws/AmazonRDS/current/${region}/index.json`)}`;

        try {
            // Fetch EC2 and RDS pricing in parallel
            const [ec2Response, rdsResponse] = await Promise.all([
                fetch(ec2Url),
                fetch(rdsUrl)
            ]);

            if (!ec2Response.ok || !rdsResponse.ok) {
                throw new Error("Không thể tải bảng giá từ API AWS");
            }

            const ec2Data = await ec2Response.json();
            const rdsData = await rdsResponse.json();

            let ec2UpdatedCount = 0;
            let rdsUpdatedCount = 0;

            // Process EC2 instances
            if (ec2Data.products && ec2Data.terms && ec2Data.terms.OnDemand) {
                const products = ec2Data.products;
                const onDemand = ec2Data.terms.OnDemand;

                Object.keys(products).forEach(sku => {
                    const prod = products[sku];
                    const attrs = prod.attributes || {};
                    
                    if (
                        prod.productFamily === "Compute Instance" &&
                        attrs.operatingSystem === "Linux" &&
                        attrs.tenancy === "Shared" &&
                        attrs.preInstalledSw === "NA"
                    ) {
                        const instanceType = attrs.instanceType;
                        const catalogItem = INSTANCE_CATALOG.ec2.find(item => item.name === instanceType);
                        if (catalogItem) {
                            const skuTerms = onDemand[sku];
                            if (skuTerms) {
                                const offerKey = Object.keys(skuTerms)[0];
                                if (offerKey) {
                                    const priceDimensions = skuTerms[offerKey].priceDimensions;
                                    if (priceDimensions) {
                                        const dimensionKey = Object.keys(priceDimensions)[0];
                                        if (dimensionKey) {
                                            const pricePerUnit = priceDimensions[dimensionKey].pricePerUnit;
                                            if (pricePerUnit && pricePerUnit.USD) {
                                                const rate = parseFloat(pricePerUnit.USD);
                                                if (rate > 0) {
                                                    catalogItem.rates[region] = rate;
                                                    ec2UpdatedCount++;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                });
            }

            // Process RDS Postgres Single-AZ instances
            if (rdsData.products && rdsData.terms && rdsData.terms.OnDemand) {
                const products = rdsData.products;
                const onDemand = rdsData.terms.OnDemand;

                Object.keys(products).forEach(sku => {
                    const prod = products[sku];
                    const attrs = prod.attributes || {};

                    if (
                        prod.productFamily === "Database Instance" &&
                        attrs.databaseEngine === "PostgreSQL" &&
                        attrs.deploymentOption === "Single-AZ"
                    ) {
                        const instanceType = attrs.instanceType;
                        const catalogItem = INSTANCE_CATALOG.rds.find(item => item.name === instanceType);
                        if (catalogItem) {
                            const skuTerms = onDemand[sku];
                            if (skuTerms) {
                                const offerKey = Object.keys(skuTerms)[0];
                                if (offerKey) {
                                    const priceDimensions = skuTerms[offerKey].priceDimensions;
                                    if (priceDimensions) {
                                        const dimensionKey = Object.keys(priceDimensions)[0];
                                        if (dimensionKey) {
                                            const pricePerUnit = priceDimensions[dimensionKey].pricePerUnit;
                                            if (pricePerUnit && pricePerUnit.USD) {
                                                const rate = parseFloat(pricePerUnit.USD);
                                                if (rate > 0) {
                                                    catalogItem.rates[region] = rate;
                                                    rdsUpdatedCount++;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                });
            }

            console.log(`Đồng bộ thành công! Đã cập nhật ${ec2UpdatedCount} giá EC2 và ${rdsUpdatedCount} giá RDS.`);

            // Trigger Recalculate
            recalculateAll();

            // Provide visual success feedback
            btnSync.innerHTML = `<i data-lucide="check-circle"></i> Đồng bộ thành công`;
            btnSync.classList.add('btn-success');
            lucide.createIcons();

            setTimeout(() => {
                btnSync.removeAttribute('disabled');
                btnSync.classList.remove('btn-success');
                btnSync.innerHTML = originalHTML;
                lucide.createIcons();
            }, 3000);

            alert(`Đồng bộ giá trực tiếp thành công cho vùng ${region}!\n- Đã cập nhật ${ec2UpdatedCount} dòng máy chủ EC2.\n- Đã cập nhật ${rdsUpdatedCount} dòng cơ sở dữ liệu RDS PostgreSQL.`);

        } catch (error) {
            console.error("Lỗi đồng bộ giá trực tiếp:", error);
            btnSync.removeAttribute('disabled');
            btnSync.innerHTML = originalHTML;
            lucide.createIcons();

            alert(`Không thể kết nối đến API AWS trực tuyến hoặc bị chặn CORS.\nỨng dụng đã tự động kích hoạt Bảng giá Offline tối ưu chất lượng cao cho vùng ${region}. Bạn vẫn có thể tiếp tục sử dụng bình thường!`);
        }
    };

    // 11. TRY SAMPLE DATA (Loads standard Vietnamese portfolio immediately)
    const loadSampleData = () => {
        const mockRows = [
            { "STT": 1, "Mã PM": "PM-CBS", "Tên PM": "Core Banking CBS System", "Đơn vị đầu mối nghiệp vụ": "Khối Dịch vụ Tài chính", "Chi phí hạ tầng được phân bổ 1 năm": "750,000,000", "Tên cấu phần": "Web Application Front-end", "CPU": 4, "RAM (GB)": 8, "Storage (GB)": 150 },
            { "STT": 2, "Mã PM": "PM-CBS", "Tên PM": "Core Banking CBS System", "Đơn vị đầu mối nghiệp vụ": "Khối Dịch vụ Tài chính", "Chi phí hạ tầng được phân bổ 1 năm": "750,000,000", "Tên cấu phần": "Database Engine Postgres Active", "CPU": 16, "RAM (GB)": 64, "Storage (GB)": 1500 },
            { "STT": 3, "Mã PM": "PM-CBS", "Tên PM": "Core Banking CBS System", "Đơn vị đầu mối nghiệp vụ": "Khối Dịch vụ Tài chính", "Chi phí hạ tầng được phân bổ 1 năm": "750,000,000", "Tên cấu phần": "Database Engine Postgres Standby", "CPU": 16, "RAM (GB)": 64, "Storage (GB)": 1500 },
            { "STT": 4, "Mã PM": "PM-CMS", "Tên PM": "Cổng thông tin Điện tử Khách hàng CMS", "Đơn vị đầu mối nghiệp vụ": "Phòng Marketing & CSKH", "Chi phí hạ tầng được phân bổ 1 năm": "140,000,000", "Tên cấu phần": "CMS Node Core API", "CPU": 2, "RAM (GB)": 4, "Storage (GB)": 50 },
            { "STT": 5, "Mã PM": "PM-CMS", "Tên PM": "Cổng thông tin Điện tử Khách hàng CMS", "Đơn vị đầu mối nghiệp vụ": "Phòng Marketing & CSKH", "Chi phí hạ tầng được phân bổ 1 năm": "140,000,000", "Tên cấu phần": "Database CMS User Postgres SQL", "CPU": 4, "RAM (GB)": 16, "Storage (GB)": 200 },
            { "STT": 6, "Mã PM": "PM-HRM", "Tên PM": "Hệ thống Quản lý Nhân sự HRM Pro", "Đơn vị đầu mối nghiệp vụ": "Khối Tổ chức Nhân sự", "Chi phí hạ tầng được phân bổ 1 năm": "80,000,000", "Tên cấu phần": "HRM Web Console API", "CPU": 2, "RAM (GB)": 4, "Storage (GB)": 30 },
            { "STT": 7, "Mã PM": "PM-HRM", "Tên PM": "Hệ thống Quản lý Nhân sự HRM Pro", "Đơn vị đầu mối nghiệp vụ": "Khối Tổ chức Nhân sự", "Chi phí hạ tầng được phân bổ 1 năm": "80,000,000", "Tên cấu phần": "Database HRM SQLite-Postgres", "CPU": 2, "RAM (GB)": 8, "Storage (GB)": 80 },
            { "STT": 8, "Mã PM": "PM-MB", "Tên PM": "Ứng dụng di động m-Banking iOS/Android", "Đơn vị đầu mối nghiệp vụ": "Khối Ngân hàng Số", "Chi phí hạ tầng được phân bổ 1 năm": "180,000,000", "Tên cấu phần": "API Gateway Router Container", "CPU": 4, "RAM (GB)": 8, "Storage (GB)": 40 },
            { "STT": 9, "Mã PM": "PM-MB", "Tên PM": "Ứng dụng di động m-Banking iOS/Android", "Đơn vị đầu mối nghiệp vụ": "Khối Ngân hàng Số", "Chi phí hạ tầng được phân bổ 1 năm": "180,000,000", "Tên cấu phần": "Microservice Transaction Core", "CPU": 8, "RAM (GB)": 16, "Storage (GB)": 150 },
            { "STT": 10, "Mã PM": "PM-BI", "Tên PM": "Phân tích Kinh doanh BI & Warehouse", "Đơn vị đầu mối nghiệp vụ": "Phòng Quản trị Dữ liệu", "Chi phí hạ tầng được phân bổ 1 năm": "1,100,000,000", "Tên cấu phần": "Kafka Event Streams", "CPU": 8, "RAM (GB)": 32, "Storage (GB)": 400 },
            { "STT": 11, "Mã PM": "PM-BI", "Tên PM": "Phân tích Kinh doanh BI & Warehouse", "Đơn vị đầu mối nghiệp vụ": "Phòng Quản trị Dữ liệu", "Chi phí hạ tầng được phân bổ 1 năm": "1,100,000,000", "Tên cấu phần": "Analytics DB Postgres Node", "CPU": 32, "RAM (GB)": 128, "Storage (GB)": 3500 }
        ];

        parsePortfolioFromRows(mockRows);

        // UI state badge update
        elFileName.innerText = "Dữ liệu thử nghiệm mẫu (11 Cấu phần)";
        elFileBadge.style.display = 'flex';
        elDropzone.style.display = 'none';

        // Trigger Cost Calculation
        recalculateAll();
    };

    // Remove file handler
    const removeFile = () => {
        parsedPortfolio = [];
        processedPortfolio = [];
        
        // Reset file elements
        elFileInput.value = '';
        elFileBadge.style.display = 'none';
        elDropzone.style.display = 'block';

        // Disable Export/Print
        elBtnExportExcel.setAttribute('disabled', 'true');
        elBtnPrintReport.setAttribute('disabled', 'true');

        // Reset Stats
        elStatTotalApps.innerText = "0";
        elStatTotalComponents.innerText = "0";
        elStatOnPremCost.innerText = "0đ";
        elStatAwsCost.innerText = "0đ";
        elStatAwsCostIcon.className = "stat-icon-wrapper neon-green";
        elStatAwsCostIcon.innerHTML = `<i data-lucide="trending-down"></i>`;

        // Clear Table
        elTableBody.innerHTML = `
            <tr>
                <td colspan="10" class="empty-table-state">
                    <div class="empty-state-content">
                        <i data-lucide="file-warning"></i>
                        <h4>Chưa có dữ liệu phân tích</h4>
                        <p>Vui lòng kéo thả file Excel vào bảng điều khiển phía trên hoặc bấm nút "Thử dữ liệu mẫu" để bắt đầu trải nghiệm.</p>
                    </div>
                </td>
            </tr>
        `;

        // Destroy charts
        if (onPremVsAwsChart) onPremVsAwsChart.destroy();
        if (awsBreakdownChart) awsBreakdownChart.destroy();
        onPremVsAwsChart = null;
        awsBreakdownChart = null;

        lucide.createIcons();
    };

    // =========================================================================
    // 11b. DYNAMIC PRICING ENGINE & LOCALSTORAGE STORAGE SYNCHRONIZATION
    // =========================================================================

    const loadPricingFromStorage = () => {
        try {
            const savedCatalog = localStorage.getItem('aws_estimator_catalog');
            const savedExchange = localStorage.getItem('aws_estimator_usd_vnd');
            
            if (savedCatalog) {
                INSTANCE_CATALOG = JSON.parse(savedCatalog);
                // Highly robust fallback for old or intermediate localStorage states
                if (!INSTANCE_CATALOG.dms || !Array.isArray(INSTANCE_CATALOG.dms) || INSTANCE_CATALOG.dms.length === 0) {
                    INSTANCE_CATALOG.dms = JSON.parse(JSON.stringify(DEFAULT_INSTANCE_CATALOG.dms));
                }
                if (!INSTANCE_CATALOG.storage) {
                    INSTANCE_CATALOG.storage = JSON.parse(JSON.stringify(DEFAULT_INSTANCE_CATALOG.storage));
                } else if (!INSTANCE_CATALOG.storage.dms) {
                    INSTANCE_CATALOG.storage.dms = JSON.parse(JSON.stringify(DEFAULT_INSTANCE_CATALOG.storage.dms));
                }
            }
            if (savedExchange) {
                USD_TO_VND = parseInt(savedExchange, 10) || 26000;
            }
        } catch (e) {
            console.error("Failed to load pricing from localStorage:", e);
        }
    };

    const savePricingToStorage = () => {
        try {
            localStorage.setItem('aws_estimator_catalog', JSON.stringify(INSTANCE_CATALOG));
            localStorage.setItem('aws_estimator_usd_vnd', USD_TO_VND.toString());
        } catch (e) {
            console.error("Failed to save pricing to localStorage:", e);
        }
    };

    const renderPricingTables = () => {
        const tBodyRds = document.getElementById('table-pricing-rds');
        const tBodyEc2 = document.getElementById('table-pricing-ec2');
        const tBodyMsk = document.getElementById('table-pricing-msk');
        const tBodyDms = document.getElementById('table-pricing-dms');

        if (!tBodyRds || !tBodyEc2 || !tBodyMsk || !tBodyDms) return;

        // Render RDS
        tBodyRds.innerHTML = INSTANCE_CATALOG.rds.map((inst, index) => {
            const monthlyRate = (inst.rates['us-east-1'] * 730).toFixed(2);
            return `
                <tr data-service="rds" data-index="${index}">
                    <td><strong class="instance-code">${inst.name}</strong></td>
                    <td><input type="number" class="table-inline-input val-cpu" value="${inst.cpu}" min="1" max="256"></td>
                    <td><input type="number" class="table-inline-input val-ram" value="${inst.ram}" min="1" max="1024"></td>
                    <td><input type="number" class="table-inline-input val-rate" value="${monthlyRate}" min="0" step="0.01"></td>
                    <td style="text-align: center;">
                        <button class="btn-row-delete" onclick="window.deleteCatalogRow('rds', ${index})" title="Xóa dòng máy này">
                            <i data-lucide="trash-2"></i>
                        </button>
                    </td>
                </tr>
            `;
        }).join('');

        // Render EC2
        tBodyEc2.innerHTML = INSTANCE_CATALOG.ec2.map((inst, index) => {
            const monthlyRate = (inst.rates['us-east-1'] * 730).toFixed(2);
            return `
                <tr data-service="ec2" data-index="${index}">
                    <td><strong class="instance-code">${inst.name}</strong></td>
                    <td><input type="number" class="table-inline-input val-cpu" value="${inst.cpu}" min="1" max="256"></td>
                    <td><input type="number" class="table-inline-input val-ram" value="${inst.ram}" min="1" max="1024"></td>
                    <td><input type="number" class="table-inline-input val-rate" value="${monthlyRate}" min="0" step="0.01"></td>
                    <td style="text-align: center;">
                        <button class="btn-row-delete" onclick="window.deleteCatalogRow('ec2', ${index})" title="Xóa dòng máy này">
                            <i data-lucide="trash-2"></i>
                        </button>
                    </td>
                </tr>
            `;
        }).join('');

        // Render MSK
        tBodyMsk.innerHTML = INSTANCE_CATALOG.msk.map((inst, index) => {
            const monthlyRate = (inst.rates['us-east-1'] * 730).toFixed(2);
            return `
                <tr data-service="msk" data-index="${index}">
                    <td><strong class="instance-code">${inst.name}</strong></td>
                    <td><input type="number" class="table-inline-input val-cpu" value="${inst.cpu}" min="1" max="256"></td>
                    <td><input type="number" class="table-inline-input val-ram" value="${inst.ram}" min="1" max="1024"></td>
                    <td><input type="number" class="table-inline-input val-rate" value="${monthlyRate}" min="0" step="0.01"></td>
                    <td style="text-align: center;">
                        <button class="btn-row-delete" onclick="window.deleteCatalogRow('msk', ${index})" title="Xóa dòng máy này">
                            <i data-lucide="trash-2"></i>
                        </button>
                    </td>
                </tr>
            `;
        }).join('');

        // Render DMS
        tBodyDms.innerHTML = INSTANCE_CATALOG.dms.map((inst, index) => {
            const monthlyRate = (inst.rates['us-east-1'] * 730).toFixed(2);
            return `
                <tr data-service="dms" data-index="${index}">
                    <td><strong class="instance-code">${inst.name}</strong></td>
                    <td><input type="number" class="table-inline-input val-cpu" value="${inst.cpu}" min="1" max="256"></td>
                    <td><input type="number" class="table-inline-input val-ram" value="${inst.ram}" min="1" max="1024"></td>
                    <td><input type="number" class="table-inline-input val-rate" value="${monthlyRate}" min="0" step="0.01"></td>
                    <td><input type="text" class="table-inline-input val-note" value="${inst.note || ''}" style="width: 100%;"></td>
                    <td style="text-align: center;">
                        <button class="btn-row-delete" onclick="window.deleteCatalogRow('dms', ${index})" title="Xóa dòng máy này">
                            <i data-lucide="trash-2"></i>
                        </button>
                    </td>
                </tr>
            `;
        }).join('');

        // Recreate icons inside dynamically generated tables
        lucide.createIcons();
    };

    window.deleteCatalogRow = (serviceType, index) => {
        if (confirm(`Bạn chắc chắn muốn xóa dòng máy này khỏi danh mục ${serviceType.toUpperCase()}?`)) {
            INSTANCE_CATALOG[serviceType].splice(index, 1);
            renderPricingTables();
        }
    };

    const savePricingUI = () => {
        // Read exchange rate
        const elUsdVnd = document.getElementById('pricing-usd-vnd');
        const elStorageEbs = document.getElementById('pricing-storage-ebs');
        const elStorageRds = document.getElementById('pricing-storage-rds');
        const elStorageMsk = document.getElementById('pricing-storage-msk');
        const elDto = document.getElementById('pricing-data-transfer-out');

        if (elUsdVnd) USD_TO_VND = parseInt(elUsdVnd.value, 10) || 26000;

        // Update Storage rates
        const ebsVal = parseFloat(elStorageEbs?.value || "9.60") / 100;
        const rdsVal = parseFloat(elStorageRds?.value || "13.80") / 100;
        const mskVal = parseFloat(elStorageMsk?.value || "10.00") / 100;
        const dtoVal = parseFloat(elDto?.value || "0.12");

        INSTANCE_CATALOG.dataTransferOutRate = dtoVal;

        const regions = ['us-east-1', 'ap-southeast-1', 'ap-northeast-1', 'eu-central-1'];
        regions.forEach(r => {
            INSTANCE_CATALOG.storage.ebs[r] = ebsVal;
            INSTANCE_CATALOG.storage.rds[r] = rdsVal;
            INSTANCE_CATALOG.storage.msk[r] = mskVal;
            INSTANCE_CATALOG.storage.dms[r] = ebsVal;
        });

        // Parse tables
        const parseTable = (tbodyId, serviceType) => {
            const rows = document.querySelectorAll(`#${tbodyId} tr`);
            rows.forEach((row, idx) => {
                const name = row.querySelector('.instance-code').innerText;
                const cpu = parseInt(row.querySelector('.val-cpu').value, 10) || 2;
                const ram = parseInt(row.querySelector('.val-ram').value, 10) || 8;
                const monthlyRate = parseFloat(row.querySelector('.val-rate').value) || 100;
                const elNote = row.querySelector('.val-note');

                // Find the item in our array and update it
                const item = INSTANCE_CATALOG[serviceType][idx];
                if (item) {
                    item.cpu = cpu;
                    item.ram = ram;
                    regions.forEach(r => {
                        item.rates[r] = monthlyRate / 730;
                    });
                    if (elNote) {
                        item.note = elNote.value;
                    }
                }
            });
        };

        parseTable('table-pricing-rds', 'rds');
        parseTable('table-pricing-ec2', 'ec2');
        parseTable('table-pricing-msk', 'msk');
        parseTable('table-pricing-dms', 'dms');

        // Save to storage
        savePricingToStorage();

        // Recalculate everything
        recalculateAll();

        alert("Đã lưu và áp dụng bảng giá động thành công! Toàn bộ tính toán và biểu đồ đã được cập nhật.");
    };

    const resetPricingUI = () => {
        if (confirm("Bạn có chắc chắn muốn khôi phục bảng giá AWS về cấu hình mặc định? Toàn bộ tùy chỉnh và máy chủ tự thêm sẽ bị xóa.")) {
            localStorage.removeItem('aws_estimator_catalog');
            localStorage.removeItem('aws_estimator_usd_vnd');

            INSTANCE_CATALOG = JSON.parse(JSON.stringify(DEFAULT_INSTANCE_CATALOG));
            USD_TO_VND = 26000;

            // Populate form fields
            const elUsdVnd = document.getElementById('pricing-usd-vnd');
            const elStorageEbs = document.getElementById('pricing-storage-ebs');
            const elStorageRds = document.getElementById('pricing-storage-rds');
            const elStorageMsk = document.getElementById('pricing-storage-msk');
            const elDto = document.getElementById('pricing-data-transfer-out');

            if (elUsdVnd) elUsdVnd.value = "26000";
            if (elStorageEbs) elStorageEbs.value = "9.60";
            if (elStorageRds) elStorageRds.value = "13.80";
            if (elStorageMsk) elStorageMsk.value = "10.00";
            if (elDto) elDto.value = "0.12";

            renderPricingTables();
            recalculateAll();

            alert("Khôi phục mặc định thành công!");
        }
    };

    const handleAddInstance = (e) => {
        e.preventDefault();
        const serviceType = document.getElementById('add-service-type').value;
        let name = document.getElementById('add-inst-name').value.trim();
        const cpu = parseInt(document.getElementById('add-inst-cpu').value, 10);
        const ram = parseInt(document.getElementById('add-inst-ram').value, 10);
        const monthlyRate = parseFloat(document.getElementById('add-inst-rate').value);
        const note = serviceType === 'dms' ? document.getElementById('add-inst-note').value.trim() : undefined;

        if (!name || isNaN(cpu) || isNaN(ram) || isNaN(monthlyRate)) {
            alert("Vui lòng điền đầy đủ và chính xác thông tin cấu hình máy chủ.");
            return;
        }

        // Format name to match naming standards of database engine
        if (serviceType === 'rds' && !name.startsWith('db.')) {
            name = `db.${name}`;
        } else if (serviceType === 'msk' && !name.startsWith('kafka.')) {
            name = `kafka.${name}`;
        }

        // Create rates dictionary
        const rates = {};
        const regions = ['us-east-1', 'ap-southeast-1', 'ap-northeast-1', 'eu-central-1'];
        regions.forEach(r => {
            rates[r] = monthlyRate / 730;
        });

        // Add to instance catalog
        const newInst = { name, cpu, ram, rates };
        if (note !== undefined) {
            newInst.note = note;
        }
        INSTANCE_CATALOG[serviceType].push(newInst);

        // Sort catalog ascending/descending so list stays clean
        INSTANCE_CATALOG[serviceType].sort((a, b) => a.cpu - b.cpu || a.ram - b.ram);

        // Reset inputs
        document.getElementById('add-inst-name').value = '';
        document.getElementById('add-inst-cpu').value = '2';
        document.getElementById('add-inst-ram').value = '8';
        document.getElementById('add-inst-rate').value = '100';
        const elNote = document.getElementById('add-inst-note');
        if (elNote) elNote.value = '';
        const elNoteWrapper = document.getElementById('add-inst-note-wrapper');
        if (elNoteWrapper) elNoteWrapper.style.display = 'none';
        document.getElementById('add-service-type').value = 'ec2';

        // Render table
        renderPricingTables();

        alert(`Đã thêm máy chủ ${name} (${cpu} vCPU, ${ram} GB RAM) vào danh mục ${serviceType.toUpperCase()} thành công! Bấm "Lưu & Áp dụng" để kích hoạt tính toán.`);
    };

    const initTabSwitching = () => {
        const tabs = document.querySelectorAll('.app-tabs .tab-btn');
        tabs.forEach(btn => {
            btn.addEventListener('click', () => {
                // Remove active classes
                tabs.forEach(t => t.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.style.display = 'none');

                // Add active to current click
                btn.classList.add('active');
                const targetTabId = btn.getAttribute('data-tab');
                const targetContent = document.getElementById(targetTabId);
                if (targetContent) {
                    targetContent.style.display = 'block';
                }

                // If switching to estimator, refresh charts just in case
                if (targetTabId === 'tab-estimator') {
                    if (processedPortfolio.length > 0) {
                        recalculateAll();
                    }
                }
                if (targetTabId === 'tab-parameters') {
                    renderGroupParametersTable();
                    renderComplexityTable();
                }
            });
        });
    };

    const initPricingUI = () => {
        // Load stored state
        loadPricingFromStorage();

        // Populate left-panel inputs from INSTANCE_CATALOG and state
        const elUsdVnd = document.getElementById('pricing-usd-vnd');
        const elStorageEbs = document.getElementById('pricing-storage-ebs');
        const elStorageRds = document.getElementById('pricing-storage-rds');
        const elStorageMsk = document.getElementById('pricing-storage-msk');
        const elDto = document.getElementById('pricing-data-transfer-out');

        if (elUsdVnd) elUsdVnd.value = USD_TO_VND;
        if (elStorageEbs) elStorageEbs.value = (INSTANCE_CATALOG.storage.ebs['us-east-1'] * 100).toFixed(2);
        if (elStorageRds) elStorageRds.value = (INSTANCE_CATALOG.storage.rds['us-east-1'] * 100).toFixed(2);
        if (elStorageMsk) elStorageMsk.value = (INSTANCE_CATALOG.storage.msk['us-east-1'] * 100).toFixed(2);
        if (elDto) {
            elDto.value = (INSTANCE_CATALOG.dataTransferOutRate !== undefined ? INSTANCE_CATALOG.dataTransferOutRate : 0.12).toFixed(2);
        }

        // Render tables
        renderPricingTables();
    };

    const loadGroupParameters = () => {
        try {
            const savedGroups = localStorage.getItem('aws_estimator_groups');
            if (savedGroups) {
                GROUP_INFO = JSON.parse(savedGroups);
                // Ensure dbChangeRate is set on all items, if missing
                Object.keys(DEFAULT_GROUP_INFO).forEach(key => {
                    if (!GROUP_INFO[key]) {
                        GROUP_INFO[key] = JSON.parse(JSON.stringify(DEFAULT_GROUP_INFO[key]));
                    }
                    if (GROUP_INFO[key].dbChangeRate === undefined) {
                        GROUP_INFO[key].dbChangeRate = DEFAULT_GROUP_INFO[key].dbChangeRate;
                    }
                });
            } else {
                GROUP_INFO = JSON.parse(JSON.stringify(DEFAULT_GROUP_INFO));
            }
        } catch (e) {
            console.error("Failed to load group parameters from localStorage:", e);
            GROUP_INFO = JSON.parse(JSON.stringify(DEFAULT_GROUP_INFO));
        }
    };

    const saveGroupParameters = () => {
        try {
            localStorage.setItem('aws_estimator_groups', JSON.stringify(GROUP_INFO));
        } catch (e) {
            console.error("Failed to save group parameters to localStorage:", e);
        }
    };

    const loadComplexityParameters = () => {
        try {
            const savedComplexity = localStorage.getItem('aws_estimator_complexity');
            if (savedComplexity) {
                COMPLEXITY_INFO = JSON.parse(savedComplexity);
            } else {
                COMPLEXITY_INFO = JSON.parse(JSON.stringify(DEFAULT_COMPLEXITY_INFO));
            }

            const savedMandayRate = localStorage.getItem('aws_estimator_manday_rate');
            if (savedMandayRate) {
                MANDAY_RATE = parseInt(savedMandayRate) || DEFAULT_MANDAY_RATE;
            } else {
                MANDAY_RATE = DEFAULT_MANDAY_RATE;
            }
        } catch (e) {
            console.error("Failed to load complexity parameters:", e);
            COMPLEXITY_INFO = JSON.parse(JSON.stringify(DEFAULT_COMPLEXITY_INFO));
            MANDAY_RATE = DEFAULT_MANDAY_RATE;
        }
    };

    const saveComplexityParameters = () => {
        try {
            localStorage.setItem('aws_estimator_complexity', JSON.stringify(COMPLEXITY_INFO));
            localStorage.setItem('aws_estimator_manday_rate', MANDAY_RATE.toString());
        } catch (e) {
            console.error("Failed to save complexity parameters:", e);
        }
    };

    const renderComplexityTable = () => {
        const tbody = document.getElementById('table-complexity-body');
        if (!tbody) return;

        tbody.innerHTML = COMPLEXITY_INFO.map((item, index) => {
            return `
                <tr data-index="${index}">
                    <td style="text-align: center;"><strong class="pm-code" style="color: var(--primary-color);">${item.level}</strong></td>
                    <td><input type="number" class="table-inline-input val-min-comp" value="${item.minComp}" min="0" style="width: 100%; text-align: center;"></td>
                    <td><input type="number" class="table-inline-input val-max-comp" value="${item.maxComp}" min="0" style="width: 100%; text-align: center;"></td>
                    <td><input type="number" class="table-inline-input val-mandays" value="${item.mandays}" min="0" style="width: 100%; text-align: center; font-weight: bold; color: var(--primary-color);"></td>
                    <td><input type="text" class="table-inline-input val-note" value="${item.note}" style="width: 100%;"></td>
                </tr>
            `;
        }).join('');

        const elMandayRate = document.getElementById('param-manday-rate');
        if (elMandayRate) {
            elMandayRate.value = MANDAY_RATE;
        }
    };

    const renderGroupParametersTable = () => {
        const tbody = document.getElementById('table-parameters-body');
        if (!tbody) return;

        tbody.innerHTML = Object.keys(GROUP_INFO).map(key => {
            const group = GROUP_INFO[key];
            return `
                <tr data-group-code="${group.code}">
                    <td style="text-align: center;"><strong class="pm-code">${group.code}</strong></td>
                    <td><input type="text" class="table-inline-input val-group-name" value="${group.name}" style="width: 100%;"></td>
                    <td><input type="text" class="table-inline-input val-stability-abbr" value="${group.stabilityAbbr}" style="width: 100%; text-align: center; font-weight: bold;"></td>
                    <td><input type="text" class="table-inline-input val-stability" value="${group.stability}" style="width: 100%;"></td>
                    <td><input type="text" class="table-inline-input val-performance-abbr" value="${group.performanceAbbr}" style="width: 100%; text-align: center; font-weight: bold;"></td>
                    <td><input type="text" class="table-inline-input val-performance" value="${group.performance}" style="width: 100%;"></td>
                    <td>
                        <div style="display: flex; align-items: center; justify-content: center; gap: 5px;">
                            <input type="number" class="table-inline-input val-db-change-rate" value="${group.dbChangeRate}" min="0" max="100" step="0.1" style="width: 80px; text-align: right; padding-right: 5px;">
                            <span style="font-weight: 600; color: var(--text-muted);">%</span>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');
    };

    const saveGroupParametersUI = () => {
        const rows = document.querySelectorAll('#table-parameters-body tr');
        rows.forEach(row => {
            const code = row.getAttribute('data-group-code');
            const name = row.querySelector('.val-group-name').value.trim();
            const stabilityAbbr = row.querySelector('.val-stability-abbr').value.trim();
            const stability = row.querySelector('.val-stability').value.trim();
            const performanceAbbr = row.querySelector('.val-performance-abbr').value.trim();
            const performance = row.querySelector('.val-performance').value.trim();
            const dbChangeRate = parseFloat(row.querySelector('.val-db-change-rate').value) || 0;

            if (GROUP_INFO[code]) {
                GROUP_INFO[code].name = name;
                GROUP_INFO[code].stabilityAbbr = stabilityAbbr;
                GROUP_INFO[code].stability = stability;
                GROUP_INFO[code].performanceAbbr = performanceAbbr;
                GROUP_INFO[code].performance = performance;
                GROUP_INFO[code].dbChangeRate = dbChangeRate;
            }
        });

        // Read complexity parameters
        const complexityRows = document.querySelectorAll('#table-complexity-body tr');
        complexityRows.forEach(row => {
            const index = parseInt(row.getAttribute('data-index'));
            const minComp = parseInt(row.querySelector('.val-min-comp').value) || 0;
            const maxComp = parseInt(row.querySelector('.val-max-comp').value) || 0;
            const mandays = parseInt(row.querySelector('.val-mandays').value) || 0;
            const note = row.querySelector('.val-note').value.trim();

            if (COMPLEXITY_INFO[index]) {
                COMPLEXITY_INFO[index].minComp = minComp;
                COMPLEXITY_INFO[index].maxComp = maxComp;
                COMPLEXITY_INFO[index].mandays = mandays;
                COMPLEXITY_INFO[index].note = note;
            }
        });

        // Read manday rate
        const elMandayRate = document.getElementById('param-manday-rate');
        if (elMandayRate) {
            MANDAY_RATE = parseInt(elMandayRate.value) || DEFAULT_MANDAY_RATE;
        }

        saveGroupParameters();
        saveComplexityParameters();
        recalculateAll();
        alert("Đã lưu và áp dụng toàn bộ tham số phân nhóm & mức độ phức tạp ứng dụng thành công!");
    };

    const resetGroupParametersUI = () => {
        if (confirm("Bạn có chắc chắn muốn khôi phục toàn bộ các tham số phân nhóm & mức độ phức tạp về cấu hình mặc định ban đầu?")) {
            localStorage.removeItem('aws_estimator_groups');
            localStorage.removeItem('aws_estimator_complexity');
            localStorage.removeItem('aws_estimator_manday_rate');
            GROUP_INFO = JSON.parse(JSON.stringify(DEFAULT_GROUP_INFO));
            loadComplexityParameters();
            renderGroupParametersTable();
            renderComplexityTable();
            recalculateAll();
            alert("Khôi phục toàn bộ tham số mặc định thành công!");
        }
    };

    const initGroupParametersUI = () => {
        loadGroupParameters();
        loadComplexityParameters();
        renderGroupParametersTable();
        renderComplexityTable();

        const btnSaveGroups = document.getElementById('btn-save-groups');
        const btnResetGroups = document.getElementById('btn-reset-groups');

        if (btnSaveGroups) {
            btnSaveGroups.addEventListener('click', saveGroupParametersUI);
        }
        if (btnResetGroups) {
            btnResetGroups.addEventListener('click', resetGroupParametersUI);
        }
    };

    // 12. EVENT BINDINGS
    
    // Dropzone drag-and-drop bindings
    elDropzone.addEventListener('click', () => elFileInput.click());
    
    elFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    elDropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        elDropzone.classList.add('dragover');
    });

    elDropzone.addEventListener('dragleave', () => {
        elDropzone.classList.remove('dragover');
    });

    elDropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        elDropzone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });

    // Remove file trigger
    elBtnRemoveFile.addEventListener('click', removeFile);

    // Global settings adjustments
    elAwsRegion.addEventListener('change', recalculateAll);
    elAwsPlan.addEventListener('change', recalculateAll);
    const elComplianceBackup = document.getElementById('compliance-realtime-backup');
    if (elComplianceBackup) {
        elComplianceBackup.addEventListener('change', recalculateAll);
    }
    const elMigrationCost = document.getElementById('compliance-migration-cost');
    if (elMigrationCost) {
        elMigrationCost.addEventListener('change', recalculateAll);
    }
    
    elAncillaryPct.addEventListener('input', (e) => {
        const val = e.target.value;
        elValAncillaryPct.innerText = `${val}%`;
        recalculateAll();
    });

    // Search bar input filter
    elSearchInput.addEventListener('input', renderTableRows);

    // Header secondary actions triggers
    elBtnDownloadTemplate.addEventListener('click', downloadTemplate);
    elBtnLoadSample.addEventListener('click', loadSampleData);
    
    // Table Export action triggers
    elBtnExportExcel.addEventListener('click', exportResultsToExcel);
    elBtnPrintReport.addEventListener('click', () => window.print());

    const elBtnSyncPricing = document.getElementById('btn-sync-pricing');
    if (elBtnSyncPricing) {
        elBtnSyncPricing.addEventListener('click', syncLivePricing);
    }

    // Pricing Management Bindings
    const elBtnSavePricing = document.getElementById('btn-save-pricing');
    if (elBtnSavePricing) {
        elBtnSavePricing.addEventListener('click', savePricingUI);
    }

    const elBtnResetPricing = document.getElementById('btn-reset-pricing');
    if (elBtnResetPricing) {
        elBtnResetPricing.addEventListener('click', resetPricingUI);
    }

    const elFormAddInstance = document.getElementById('form-add-instance');
    if (elFormAddInstance) {
        elFormAddInstance.addEventListener('submit', handleAddInstance);
    }

    const elAddServiceType = document.getElementById('add-service-type');
    if (elAddServiceType) {
        elAddServiceType.addEventListener('change', (e) => {
            const wrapper = document.getElementById('add-inst-note-wrapper');
            if (wrapper) {
                wrapper.style.display = e.target.value === 'dms' ? 'block' : 'none';
            }
        });
    }

    // ==========================================
    // 12. RESULTS TABLE SORTING BEHAVIORS
    // ==========================================
    sortPortfolio = () => {
        if (!currentSortColumn || processedPortfolio.length === 0) return;

        processedPortfolio.sort((a, b) => {
            let valA, valB;

            switch (currentSortColumn) {
                case 'stt':
                    valA = a.originalIndex;
                    valB = b.originalIndex;
                    break;
                case 'maPM':
                    valA = a.maPM.toLowerCase();
                    valB = b.maPM.toLowerCase();
                    break;
                case 'tenPM':
                    valA = a.tenPM.toLowerCase();
                    valB = b.tenPM.toLowerCase();
                    break;
                case 'owner':
                    valA = a.owner.toLowerCase();
                    valB = b.owner.toLowerCase();
                    break;
                case 'onPremCost':
                    valA = a.onPremCost;
                    valB = b.onPremCost;
                    break;
                case 'grandTotalAwsVND':
                    valA = a.grandTotalAwsVND;
                    valB = b.grandTotalAwsVND;
                    break;
                case 'deltaVND':
                    valA = a.deltaVND;
                    valB = b.deltaVND;
                    break;
                case 'deltaPct':
                    valA = a.deltaPct;
                    valB = b.deltaPct;
                    break;
                case 'phanNhom':
                    valA = (a.phanNhom || "").toLowerCase();
                    valB = (b.phanNhom || "").toLowerCase();
                    break;
                case 'stability':
                    valA = (a.stability || "").toLowerCase();
                    valB = (b.stability || "").toLowerCase();
                    break;
                case 'performance':
                    valA = (a.performance || "").toLowerCase();
                    valB = (b.performance || "").toLowerCase();
                    break;
                case 'recommend':
                    valA = a.recommend ? 1 : 0;
                    valB = b.recommend ? 1 : 0;
                    break;
                default:
                    return 0;
            }

            if (valA < valB) return currentSortDirection === 'asc' ? -1 : 1;
            if (valA > valB) return currentSortDirection === 'asc' ? 1 : -1;
            return 0;
        });
    };

    updateHeaderSortUI = () => {
        const headers = document.querySelectorAll('#results-table th.sortable');
        headers.forEach(th => {
            const colField = th.getAttribute('data-sort');
            th.classList.remove('active-sort');
            
            const icon = th.querySelector('.sort-icon');
            if (icon) {
                if (colField === currentSortColumn) {
                    th.classList.add('active-sort');
                    const iconName = currentSortDirection === 'asc' ? 'chevron-up' : 'chevron-down';
                    icon.setAttribute('data-lucide', iconName);
                } else {
                    icon.setAttribute('data-lucide', 'chevrons-up-down');
                }
            }
        });
        
        lucide.createIcons();
    };

    const initTableSorting = () => {
        const headers = document.querySelectorAll('#results-table th.sortable');
        headers.forEach(th => {
            th.addEventListener('click', () => {
                const colField = th.getAttribute('data-sort');
                if (colField === currentSortColumn) {
                    currentSortDirection = currentSortDirection === 'asc' ? 'desc' : 'asc';
                } else {
                    currentSortColumn = colField;
                    currentSortDirection = 'asc';
                }
                
                sortPortfolio();
                renderTableRows();
                updateHeaderSortUI();
            });
        });
    };

    // Initialize state & behaviors
    initPricingUI();
    initGroupParametersUI();
    initTabSwitching();
    initTableSorting();
    updateHeaderSortUI();

    // Add helper to set dates dynamically for print footer reports
    const updatePrintDateAttribute = () => {
        const resultsSection = document.querySelector('.results-section');
        const todayStr = new Date().toLocaleDateString('vi-VN', { year: 'numeric', month: 'long', day: 'numeric' });
        resultsSection.setAttribute('data-date', todayStr);
    };
    updatePrintDateAttribute();
});
