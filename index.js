const api = require('./api');
const excel = require('./excel');
const ui = require('./ui');
const config = require('./config.json');
const Table = require('cli-table3');

const fs = require('fs');
const xlsx = require('xlsx');
let chalk, inquirer;

const jeda = (minDetik, maksDetik) => {
    const durasi = (Math.floor(Math.random() * (maksDetik - minDetik + 1)) + minDetik) * 1000;
    if (durasi > 0) {
        console.log(chalk.gray(`   ⏸️  Jeda acak selama ${durasi / 1000} detik...`));
    }
    return new Promise(resolve => setTimeout(resolve, durasi));
};

async function runBuatMasterPelanggan() {
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- [Utilitas] Membuat/Memperbarui Master Pelanggan ---"));

    console.log(chalk.blue("1. Memindai seluruh data transaksi dari semua pangkalan..."));
    const { data: semuaLog } = excel.bacaSemuaSheetLog(config.filePaths.masterLogTransaksi);

    const pelangganMap = new Map();
    let totalTransaksiDibaca = 0;

    for (const pangkalan in semuaLog) {
        semuaLog[pangkalan].forEach(row => {
            if (!row.noKTP || !row.nama || row.nama === 'Akan diverifikasi' || String(row.noKTP).includes('x')) return;
            totalTransaksiDibaca++;
            const nik = String(row.noKTP);
            
            const dataPelangganSaatIni = pelangganMap.get(nik);
            const tanggalBaru = new Date(String(row.tanggal_transaksi).split(',')[0].split('/').reverse().join('-'));

            if (!dataPelangganSaatIni || !dataPelangganSaatIni.tanggal_terakhir_transaksi || dataPelangganSaatIni.tanggal_terakhir_transaksi === 'Belum Pernah Transaksi' || tanggalBaru > new Date(String(dataPelangganSaatIni.tanggal_terakhir_transaksi).split(',')[0].split('/').reverse().join('-'))) {
                pelangganMap.set(nik, {
                    noKTP: nik,
                    nama: row.nama,
                    customerTypes: row.customerTypes,
                    tanggal_terakhir_transaksi: row.tanggal_transaksi,
                });
            }
        });
    }

    if (pelangganMap.size === 0) {
        console.log(chalk.yellow("Tidak ada data valid di master log untuk diproses."));
        return;
    }

    const dataPelanggan = Array.from(pelangganMap.values());
    excel.tulisLog(config.filePaths.masterPelanggan, xlsx.utils.book_new(), "Pelanggan", dataPelanggan);

    console.log(chalk.green.bold(`\n✅ File "${config.filePaths.masterPelanggan}" berhasil dibuat/diperbarui!`));
    console.log(`   > Diproses ${totalTransaksiDibaca} baris transaksi.`);
    console.log(`   > Ditemukan dan disimpan ${dataPelanggan.length} pelanggan unik.`);
    console.log(`⏱️  Waktu Proses: ${ui.formatWaktuProses(Date.now() - startTime)}`);
}

const parseTanggalTransaksi = (tanggalString) => {
    if (!tanggalString || typeof tanggalString !== 'string') return null;
    
    // Ambil bagian tanggal saja (sebelum spasi atau koma pertama)
    const datePart = tanggalString.split(' ')[0].split(',')[0];
    const parts = datePart.split('/');
    
    // Pastikan formatnya DD/MM/YYYY
    if (parts.length !== 3) return null; 
    
    // Gabungkan kembali ke format YYYY-MM-DD yang aman untuk semua sistem
    const formattedDateString = `${parts[2]}-${parts[1]}-${parts[0]}`;
    const dateObj = new Date(formattedDateString);

    // Kembalikan null jika hasilnya tanggal tidak valid
    return isNaN(dateObj) ? null : dateObj;
};

async function backupMasterLog() {
    const filePath = config.filePaths.masterLogTransaksi;
    if (!fs.existsSync(filePath)) return;

    const timestamp = new Date().toISOString().replace(/:/g, '-').slice(0, 19);
    const backupPath = filePath.replace('.xlsx', `_backup_${timestamp}.xlsx`);
    fs.copyFileSync(filePath, backupPath);
    console.log(chalk.gray(`   > Backup master log dibuat di: ${backupPath}`));
}

async function runTambahPelangganBaru(token) {
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- [Manajemen] Menambahkan Pelanggan Baru dari File ---"));

    const dataPelangganBaru = excel.bacaFile(config.filePaths.inputFilePelangganBaru);
    if (!dataPelangganBaru) {
        console.log(chalk.red(`   Buat file "${config.filePaths.inputFilePelangganBaru}" dengan satu kolom "noKTP".`));
        return;
    }

    const dataPelanggan = excel.bacaFile(config.filePaths.masterPelanggan) || [];
    const pelangganSet = new Set(dataPelanggan.map(p => String(p.noKTP)));

    const summary = { 'Baru Ditambahkan': 0, 'Sudah Ada (Dilewati)': 0, 'Verifikasi Gagal': 0 };
    const progressBar = ui.buatProgressBar('Memproses NIK Baru');
    progressBar.start(dataPelangganBaru.length, 0);

    let adaDataBaru = false;

    for (const [index, user] of dataPelangganBaru.entries()) {
        progressBar.increment();
        const nik = String(user.noKTP);
        if (!nik || !/^\d{16}$/.test(nik)) {
            summary['Verifikasi Gagal']++;
            continue;
        }

        if (pelangganSet.has(nik)) {
            summary['Sudah Ada (Dilewati)']++;
            continue;
        }

        const verificationData = await api.getVerificationData(nik, token);
        if (verificationData) {
            const quota = await api.getQuota(nik, verificationData, token);
            
            const custType = verificationData.customerTypes?.[0]?.name || 'N/A';
            dataPelanggan.push({
                noKTP: nik,
                nama: verificationData.name,
                customerTypes: custType,
                tanggal_terakhir_transaksi: 'Belum Pernah Transaksi',
                daily: quota?.daily ?? 0,
                monthly: quota?.monthly ?? 0,
                family: quota?.family ?? 'N/A',
                all: quota?.all ?? 0
            });
            pelangganSet.add(nik);
            summary['Baru Ditambahkan']++;
            adaDataBaru = true;
        } else {
            summary['Verifikasi Gagal']++;
        }

        if (index < dataPelangganBaru.length - 1) {
            await jeda(config.jeda.cekKuota.minDetik, config.jeda.cekKuota.maksDetik);
        }
    }
    progressBar.stop();

    if (adaDataBaru) {
        excel.tulisLog(config.filePaths.masterPelanggan, xlsx.utils.book_new(), "Pelanggan", dataPelanggan);
        console.log(chalk.green.bold("\n\n✅ Master Pelanggan berhasil diperbarui."));
    } else {
        console.log(chalk.yellow("\n\nTidak ada pelanggan baru yang ditambahkan."));
    }

    ui.tampilkanTabelRingkasan(summary, "Ringkasan Penambahan Pelanggan");
    console.log(`⏱️  Waktu Proses: ${ui.formatWaktuProses(Date.now() - startTime)}`);
}                                                                                                                                  

async function runTransaksiLangsung(token, profile) {
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- [Transaksi Langsung] Input On-The-Spot ---"));
    
    // Variabel untuk mengontrol perulangan
    let lakukanLagi = true;

    while (lakukanLagi) {
        try {
            const pathMasterPelanggan = config.filePaths.masterPelanggan;
            let dataPelanggan = excel.bacaFile(pathMasterPelanggan) || [];
            const pelangganSet = new Set(dataPelanggan.map(p => String(p.noKTP)));
            
            const { nik, quantity } = await ui.promptTransaksiLangsung();
            
            console.log(chalk.blue("\nMemproses transaksi..."));
            
            const productInfo = await api.getProducts(token);
            if (quantity > productInfo.stockAvailable) {
                throw new Error("Stok tidak cukup untuk transaksi ini.");
            }

            const sheetName = profile.storeName.replace(/[\\/*?:"\[\]]/g, '').substring(0, 31);
            const { data: historicalData, workbook } = excel.bacaLog(config.filePaths.masterLogTransaksi, sheetName);
            
            const currentMonth = new Date().getMonth();
            const usageMap = new Map();
            historicalData
                .filter(row => row.status?.startsWith('Sukses') && new Date(String(row.tanggal_transaksi).split(',')[0].split('/').reverse().join('-')).getMonth() === currentMonth)
                .forEach(row => usageMap.set(String(row.noKTP), (usageMap.get(String(row.noKTP)) || 0) + (row.quantity || 1)));

            const verificationData = await api.getVerificationData(nik, token);
            if (!verificationData) throw new Error("Verifikasi NIK Gagal / NIK tidak terdaftar");

            const customerTypeInfo = verificationData.customerTypes?.[0];
            const custType = customerTypeInfo?.name || 'N/A';
            const sourceTypeIdValue = customerTypeInfo?.sourceTypeId;
            if (!sourceTypeIdValue) throw new Error(`sourceTypeId tidak ditemukan untuk NIK ${nik}`);
            
            const monthlyLimit = (custType === 'Usaha Mikro') ? config.aturanBisnis.batasUsahaMikro : config.aturanBisnis.batasPerPangkalan;

            if ((usageMap.get(nik) || 0) + quantity > monthlyLimit) {
                throw new Error(`Batas ${monthlyLimit} transaksi/bulan untuk pelanggan ini akan terlampaui.`);
            }
            
            const payload = {
                products: [{ productId: productInfo.productId, quantity }], token: verificationData.token,
                subsidi: { nik, familyIdEncrypted: verificationData.familyIdEncrypted, category: custType, nama: verificationData.name, channelInject: "tnp2k", sourceTypeId: sourceTypeIdValue },
            };

            const trxData = await api.postTransaction(payload, token);
            if (!trxData.success) throw new Error(trxData.message || "TRANSACTION_INVALID");

            const tanggalTransaksi = new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'});
            const resultRow = {
                noKTP: nik, nama: verificationData.name, customerTypes: custType,
                status: `Sukses - ID: ${trxData.data.transactionId}`,
                tanggal_transaksi: tanggalTransaksi,
                pangkalan: profile.storeName, quantity,
            };
            
            historicalData.push(resultRow);
            excel.tulisLog(config.filePaths.masterLogTransaksi, workbook, sheetName, historicalData);
            
            if (!pelangganSet.has(nik)) {
                const quota = await api.getQuota(nik, verificationData, token);
                const newPelanggan = {
                    noKTP: nik, nama: verificationData.name, customerTypes: custType,
                    tanggal_terakhir_transaksi: tanggalTransaksi,
                    daily: quota?.daily ?? 0, monthly: quota?.monthly ?? 0,
                    family: quota?.family ?? 'N/A', all: quota?.all ?? 0,
                    terakhir_dicek: new Date().toISOString()
                };
                dataPelanggan.push(newPelanggan);
                console.log(chalk.blue(`   > Pelanggan baru "${verificationData.name}" ditambahkan ke master.`));
            }
            
            excel.tulisLog(config.filePaths.masterPelanggan, xlsx.utils.book_new(), "Pelanggan", dataPelanggan);
            
            console.log(chalk.green.bold(`\n✅ Transaksi untuk ${verificationData.name} berhasil!`));
            console.log(`   Stok tersisa: ${productInfo.stockAvailable - quantity}`);
            
        } catch (error) {
            console.log(chalk.red.bold(`\n❌ Gagal: ${error.message}`));
        }
        
        // Pertanyaan untuk melanjutkan perulangan
        const { continuePrompt } = await inquirer.prompt([{
            type: 'confirm',
            name: 'continuePrompt',
            message: 'Lakukan transaksi manual lagi?',
            default: true
        }]);
        lakukanLagi = continuePrompt;
    }
    
    console.log(`\n⏱️  Total Waktu Sesi Transaksi Langsung: ${ui.formatWaktuProses(Date.now() - startTime)}`);
}

async function buatRencanaTransaksi(token, profile) {
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- [Cerdas] Membuat Rencana Transaksi Harian ---"));
    
    // try {
    //     console.log(chalk.red("\n Sinkronisasi master log transaksi dengan log server..."));
    //     const { startDate, endDate } = await ui.promptPilihRentangTanggal();
    //     const tglMulai = startDate.toISOString().split('T')[0];
    //     const tglSelesai = endDate.toISOString().split('T')[0];
        
    //     console.log(chalk.blue("\n Membaca data Master Pelanggan lokal..."));
    //     // Langkah 1: Baca Master Pelanggan di awal
    //     let dataPelanggan = excel.bacaFile(config.filePaths.masterPelanggan) || [];
    //     if(dataPelanggan.length === 0) {
    //         console.log(chalk.yellow("   > Peringatan: File Master Pelanggan kosong atau tidak ditemukan."));
    //     }
        
    //     console.log(chalk.blue(`\n Mengambil laporan dari server...`));
    //     const serverCustomers = await api.getTransactionsReport(token, tglMulai, tglSelesai);
    //     if (serverCustomers.length === 0) {
    //         console.log(chalk.yellow("Tidak ada transaksi ditemukan di server pada periode ini."));
    //         return;
    //     }

    //     console.log(chalk.blue("\n Membaca dan menyinkronkan log transaksi lokal..."));
    //     const sheetName = profile.storeName.replace(/[\\/*?:"\[\]]/g, '').substring(0, 31);
    //     const { data: localData, workbook } = excel.bacaLog(config.filePaths.masterLogTransaksi, sheetName);
    //     const localTransactionIds = new Set(localData.map(row => (row.status || '').match(/ID: ([\w-]+)/)?.[1]).filter(Boolean));

    //     let transactionsAdded = 0;
    //     const progressBar = ui.buatProgressBar('Sinkronisasi');
    //     progressBar.start(serverCustomers.length, 0);

    //     for (const customer of serverCustomers) {
    //         const serverTransactions = await api.getTransactionsByCustomer(token, tglMulai, tglSelesai, customer.customerReportId);
            
    //         for (const tx of serverTransactions) {
    //             if (!localTransactionIds.has(tx.transactionId)) {
    //                 const detail = await api.getTransactionDetail(token, tx.transactionId);
                    
    //                 const maskedNik = tx.nationalityId || detail.subsidi.nik;
    //                 const customerName = detail.subsidi.nama;
    //                 const customerCategory = detail.subsidi.category;
    //                 const cleanName = (name) => name.trim().replace(/\s+/g, ' ').toLowerCase();
    //                 const serverNameClean = cleanName(customerName);
                    
    //                 let matchedCustomer = dataPelanggan.find(p => cleanName(p.nama) === serverNameClean && String(p.noKTP).slice(0, 3) === maskedNik.slice(0, 3) && String(p.noKTP).slice(-3) === maskedNik.slice(-3) );
                    
    //                 let fullNik = maskedNik;
    //                 if (matchedCustomer) { fullNik = matchedCustomer.noKTP; }
                    
    //                 if (fullNik !== maskedNik && !fullNik.includes('x')) {
    //                     console.log(chalk.cyan(`\n   > Menambahkan transaksi hilang: ID ${tx.transactionId} untuk ${customerName} (NIK: ${fullNik})`));
    //                 } else {
    //                     console.log(chalk.yellow(`\n   > Menambahkan transaksi hilang: ID ${tx.transactionId} untuk ${customerName} (NIK LENGKAP TIDAK DITEMUKAN)`));
    //                 }
                    
    //                 const newLogRow = {
    //                     noKTP: fullNik,
    //                     nama: customerName,
    //                     customerTypes: customerCategory,
    //                     status: `Sukses - ID: ${detail.transactionId}`,
    //                     tanggal_transaksi: detail.subHeader.date,
    //                     pangkalan: detail.receipt.storeName,
    //                     quantity: detail.products[0].rawValue.quantity
    //                 };
                    
    //                 localData.push(newLogRow);
    //                 transactionsAdded++;
    //             }
    //         }
    //         progressBar.increment();
    //         await jeda(0, 1);
    //     }
    //     progressBar.stop();

    //     if (transactionsAdded > 0) {
    //         console.log(chalk.green.bold(`\n\n✅ Sinkronisasi log selesai. ${transactionsAdded} transaksi baru berhasil ditambahkan.`));
    //         excel.tulisLog(config.filePaths.masterLogTransaksi, workbook, sheetName, localData);
    //     } else {
    //         console.log(chalk.green.bold("\n\n✅ Sinkronisasi log selesai. Log lokal Anda sudah sesuai dengan data server."));
    //     }
        
    //     // --- BLOK PERBAIKAN DIMULAI DI SINI ---
        
    //     console.log(chalk.blue("\n4. Memperbarui 'tanggal_terakhir_transaksi' di Master Pelanggan..."));
        
    //     // Langkah 3: Buat peta transaksi terakhir dari log yang sudah lengkap
    //     const latestTransactionMap = new Map();
    //     for(const trx of localData){
    //         if(trx.noKTP && trx.status?.startsWith('Sukses')){
    //             const nik = String(trx.noKTP);
    //             // Parsing tanggal, mengakomodasi format yang berbeda
    //             const trxDateStr = String(trx.tanggal_transaksi).split(',')[0].split(' ')[0];
    //             const parts = trxDateStr.split('/');
    //             let trxDate;
    //             if(parts.length === 3) { // Format dd/MM/yyyy
    //                 trxDate = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
    //             } else { // Coba parsing langsung jika formatnya beda
    //                 trxDate = new Date(trx.tanggal_transaksi);
    //             }

    //             if(!isNaN(trxDate.getTime())){
    //                 const existing = latestTransactionMap.get(nik);
    //                 if(!existing || trxDate > existing.date){
    //                     latestTransactionMap.set(nik, { date: trxDate, dateString: trx.tanggal_transaksi });
    //                 }
    //             }
    //         }
    //     }
        
    //     // Langkah 4: Perbarui data pelanggan di memori
    //     let masterPelangganUpdated = false;
    //     for(const pelanggan of dataPelanggan){
    //         const nik = String(pelanggan.noKTP);
    //         const latestTrx = latestTransactionMap.get(nik);
    //         if(latestTrx && pelanggan.tanggal_terakhir_transaksi !== latestTrx.dateString){
    //             pelanggan.tanggal_terakhir_transaksi = latestTrx.dateString;
    //             masterPelangganUpdated = true;
    //         }
    //     }
        
    //     // Langkah 5: Tulis kembali jika ada perubahan
    //     if(masterPelangganUpdated){
    //         excel.tulisLog(config.filePaths.masterPelanggan, xlsx.utils.book_new(), "Pelanggan", dataPelanggan);
    //         console.log(chalk.green.bold("✅ Master Pelanggan berhasil diperbarui dengan tanggal transaksi terbaru."));
    //     } else {
    //         console.log(chalk.green("   > Master Pelanggan sudah sinkron, tidak ada pembaruan tanggal transaksi."));
    //     }
    //     // --- BLOK PERBAIKAN SELESAI ---

    // } catch (error) {
    //     console.log(chalk.red.bold(`\n❌ Terjadi kesalahan saat sinkronisasi: ${error.message}`));
    // }

    const dataPelanggan = excel.bacaFile(config.filePaths.masterPelanggan);
    if (!dataPelanggan) {
        console.log(chalk.red.bold(`\n❌ File master pelanggan tidak ditemukan. Jalankan menu "Buat/Update Master Pelanggan" terlebih dahulu.`));
        return;
    }
    
    const { totalStock } = await inquirer.prompt([{
        type: 'number', name: 'totalStock',
        message: `Masukkan total stok yang akan didistribusikan untuk ${chalk.cyan(profile.storeName)}:`,
        validate: (input) => input > 0 ? true : "Jumlah stok harus lebih dari 0."
    }]);

    console.log(chalk.blue("\n1. Membaca data transaksi bulan ini..."));
    const { data: semuaLog } = excel.bacaSemuaSheetLog(config.filePaths.masterLogTransaksi);
    const globalUsageMap = new Map();
    const lastTxDateMap = new Map();
    const currentMonth = new Date().getMonth();
    
    for (const pangkalanSheet in semuaLog) {
        semuaLog[pangkalanSheet].forEach(row => {
            if (row.status?.startsWith('Sukses') && parseTanggalTransaksi(row.tanggal_transaksi)?.getMonth() === currentMonth) {
                const nik = String(row.noKTP);
                const tglBaru = parseTanggalTransaksi(row.tanggal_transaksi);
                
                // Update Peta Penggunaan Bulanan
                const pangkalanUsage = globalUsageMap.get(nik) || new Map();
                pangkalanUsage.set(pangkalanSheet, (pangkalanUsage.get(pangkalanSheet) || 0) + 1);
                globalUsageMap.set(nik, pangkalanUsage);
                
                // Update Peta Tanggal Transaksi Terakhir per Pangkalan
                const pangkalanDates = lastTxDateMap.get(nik) || new Map();
                const tglLama = pangkalanDates.get(pangkalanSheet) ? parseTanggalTransaksi(pangkalanDates.get(pangkalanSheet)) : null;
                if (!tglLama || tglBaru > tglLama) {
                    pangkalanDates.set(pangkalanSheet, row.tanggal_transaksi);
                    lastTxDateMap.set(nik, pangkalanDates);
                }
            }
        });
    }

    console.log(chalk.blue("\n2. Melakukan filter cepat berdasarkan data tersimpan..."));
    const sheetName = profile.storeName.replace(/[\\/*?:"\[\]]/g, '').substring(0, 31);
    
    let kandidatPelanggan = dataPelanggan.filter(pelanggan => {
        const nik = String(pelanggan.noKTP);
        const usageDiPangkalanIni = globalUsageMap.get(nik)?.get(sheetName) || 0;
        const limitPangkalan = (pelanggan.customerTypes === 'Usaha Mikro') ? config.aturanBisnis.batasUsahaMikro : config.aturanBisnis.batasPerPangkalan;
        
        // --- [PERBAIKAN] Logika Jarak Hari sekarang spesifik per pangkalan ---
        let lolosJarakHari = false;
        const tglTerakhirDiPangkalanIni = lastTxDateMap.get(nik)?.get(sheetName);

        if (!tglTerakhirDiPangkalanIni) {
            lolosJarakHari = true; // Lolos jika belum pernah transaksi di pangkalan ini
        } else {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const lastTxDate = parseTanggalTransaksi(tglTerakhirDiPangkalanIni);
            if(lastTxDate) {
                lastTxDate.setHours(0, 0, 0, 0);
                const selisihHari = Math.floor((today - lastTxDate) / (1000 * 60 * 60 * 24));
                if (selisihHari >= config.aturanBisnis.jarakHariMinimumTransaksi) {
                    lolosJarakHari = true;
                }
            }
        }
        
        const lolosKuotaCache = !(typeof pelanggan.monthly === 'number' && pelanggan.monthly <= 0);
        return usageDiPangkalanIni < limitPangkalan && lolosKuotaCache && lolosJarakHari;
    });
    console.log(chalk.gray(`   > Ditemukan ${kandidatPelanggan.length} kandidat awal.`));
    
    const limitVerifikasi = Math.ceil(totalStock * config.aturanBisnis.faktor_buffer_kandidat);
    const kandidatTeratas = kandidatPelanggan.slice(0, limitVerifikasi);
    console.log(chalk.blue(`\n3. Memprioritaskan dan memilih ${kandidatTeratas.length} kandidat terbaik untuk diverifikasi API...`));
    const eligibleCustomers = [];
    const progressBar = ui.buatProgressBar('Update Cerdas');
    progressBar.start(kandidatTeratas.length, 0);
    for (const [index, pelanggan] of kandidatTeratas.entries()) {
        progressBar.increment();
        const nik = String(pelanggan.noKTP);
        let isDataFresh = false;
        if (pelanggan.terakhir_dicek) {
            const hoursDiff = (new Date() - new Date(pelanggan.terakhir_dicek)) / 36e5;
            if (hoursDiff < config.aturanBisnis.masaBerlakuCacheKuotaJam) isDataFresh = true;
        }
        if (isDataFresh) {
            if (pelanggan.daily > 0) eligibleCustomers.push({ ...pelanggan, usage: globalUsageMap.get(nik)?.get(sheetName) || 0 });
        } else {
            const verificationData = await api.getVerificationData(nik, token);
            if (verificationData) {
                const quota = await api.getQuota(nik, verificationData, token);
                if (quota) {
                    pelanggan.daily = quota.daily; pelanggan.monthly = quota.monthly; pelanggan.family = quota.family; pelanggan.all = quota.all;
                    pelanggan.terakhir_dicek = new Date().toISOString();
                    if (quota.daily > 0) eligibleCustomers.push({ ...pelanggan, usage: globalUsageMap.get(nik)?.get(sheetName) || 0 });
                }
            }
            if (index < kandidatTeratas.length - 1) await jeda(config.jeda.cekKuota.minDetik, config.jeda.cekKuota.maksDetik);
        }
    }
    progressBar.stop();
    console.log(chalk.blue("\n4. Menyimpan hasil Pembaruan Cerdas ke Master Pelanggan..."));
    excel.tulisLog(config.filePaths.masterPelanggan, xlsx.utils.book_new(), "Pelanggan", dataPelanggan);

    if (eligibleCustomers.length === 0) {
        console.log(chalk.yellow("\nTidak ditemukan pelanggan yang memenuhi syarat setelah verifikasi API."));
        return;
    }

    // --- BLOK BARU: LOGIKA DUA TAHAP UNTUK KUANTITAS 1 ATAU 2 ---
    console.log(chalk.blue("\n5. Menyusun rencana transaksi dengan kuantitas dinamis (Tahap 1 & 2)..."));
    
    eligibleCustomers.sort((a, b) => {
        if (a.customerTypes === 'Usaha Mikro' && b.customerTypes !== 'Usaha Mikro') return -1;
        if (a.customerTypes !== 'Usaha Mikro' && b.customerTypes === 'Usaha Mikro') return 1;
        const tglA = a.tanggal_terakhir_transaksi === 'Belum Pernah Transaksi' ? 0 : new Date(String(a.tanggal_terakhir_transaksi).split(',')[0].split('/').reverse().join('-')).getTime();
        const tglB = b.tanggal_terakhir_transaksi === 'Belum Pernah Transaksi' ? 0 : new Date(String(b.tanggal_terakhir_transaksi).split(',')[0].split('/').reverse().join('-')).getTime();
        return tglA - tglB;
    });

    const pelangganTerpilih = [];
    let stockTerpakai = 0;

    // Tahap 1: Alokasi dasar
    for (const p of eligibleCustomers) {
        if (stockTerpakai >= totalStock) break;
        const kuantitasDefault = (p.customerTypes === 'Usaha Mikro') ? 5 : 1;
        const kuantitasDapatDiambil = Math.min(kuantitasDefault, (totalStock - stockTerpakai), p.daily, p.monthly);

        if (kuantitasDapatDiambil > 0) {
            pelangganTerpilih.push({
                noKTP: p.noKTP, nama: p.nama, customerTypes: p.customerTypes,
                quantity: kuantitasDapatDiambil, sisa_kuota_harian: p.daily,
                sisa_kuota_bulanan: p.monthly, sisa_kuota_keluarga: p.family
            });
            stockTerpakai += kuantitasDapatDiambil;
        }
    }

    // Tahap 2: Alokasi tambahan untuk Rumah Tangga jika ada sisa stok
    let sisaStok = totalStock - stockTerpakai;
    if (sisaStok > 0) {
        for (const pelanggan of pelangganTerpilih) {
            if (sisaStok <= 0) break;
            // Hanya targetkan Rumah Tangga yang kuantitasnya masih 1
            if (pelanggan.customerTypes === 'Rumah Tangga' && pelanggan.quantity === 1) {
                // Pastikan pelanggan masih punya kuota untuk tabung kedua
                if (pelanggan.sisa_kuota_harian > 1 && pelanggan.sisa_kuota_bulanan > 1) {
                    pelanggan.quantity += 1;
                    stockTerpakai += 1;
                    sisaStok -= 1;
                }
            }
        }
    }
    // --- BLOK BARU SELESAI ---

    console.log(chalk.green(`\n6. Berhasil menemukan ${eligibleCustomers.length} pelanggan, ${pelangganTerpilih.length} dipilih untuk rencana dengan total ${stockTerpakai} tabung.`));
    excel.tulisLog(config.filePaths.rencanaTransaksi, xlsx.utils.book_new(), "Rencana", pelangganTerpilih);
    
    console.log(chalk.bold.green(`\n✅ File "${config.filePaths.rencanaTransaksi}" berhasil dibuat.`));
    console.log(`⏱️  Waktu Proses: ${ui.formatWaktuProses(Date.now() - startTime)}`);
}

async function eksekusiRencanaTransaksi(token, profile) {
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- [Cerdas] Eksekusi Rencana Transaksi dengan Auto-Replace ---"));

    const dataToInput = excel.bacaFile(config.filePaths.rencanaTransaksi);
    if (!dataToInput || dataToInput.length === 0) {
        console.log(chalk.red(`File rencana transaksi tidak ditemukan atau kosong.`));
        return;
    }

    console.log(chalk.blue("1. Mempersiapkan daftar pelanggan pengganti yang valid..."));
    const dataPelanggan = excel.bacaFile(config.filePaths.masterPelanggan) || [];
    const niksDalamRencana = new Set(dataToInput.map(u => String(u.noKTP)));
    
    // --- [START] Logika Baru untuk Memfilter Kandidat Pengganti ---
    const { data: semuaLog } = excel.bacaSemuaSheetLog(config.filePaths.masterLogTransaksi);
    const globalUsageMap = new Map();
    const lastTxDateMap = new Map();
    const currentMonth = new Date().getMonth();
    const sheetName = profile.storeName.replace(/[\\/*?:"\[\]]/g, '').substring(0, 31);

    // Membangun peta penggunaan dan tanggal transaksi terakhir dari file log
    for (const pangkalanSheet in semuaLog) {
        semuaLog[pangkalanSheet].forEach(row => {
            const trxDate = parseTanggalTransaksi(row.tanggal_transaksi);
            if (row.status?.startsWith('Sukses') && trxDate) {
                const nik = String(row.noKTP);
                
                // Update Peta Penggunaan Bulanan jika transaksi terjadi di bulan ini
                if (trxDate.getMonth() === currentMonth) {
                    const pangkalanUsage = globalUsageMap.get(nik) || new Map();
                    pangkalanUsage.set(pangkalanSheet, (pangkalanUsage.get(pangkalanSheet) || 0) + (row.quantity || 1));
                    globalUsageMap.set(nik, pangkalanUsage);
                }
                
                // Update Peta Tanggal Transaksi Terakhir di setiap pangkalan
                const pangkalanDates = lastTxDateMap.get(nik) || new Map();
                const tglLama = pangkalanDates.get(pangkalanSheet) ? parseTanggalTransaksi(pangkalanDates.get(pangkalanSheet)) : null;
                if (!tglLama || trxDate > tglLama) {
                    pangkalanDates.set(pangkalanSheet, row.tanggal_transaksi);
                    lastTxDateMap.set(nik, pangkalanDates);
                }
            }
        });
    }

    // Filter dan sort kandidat pengganti berdasarkan data aktual
    let kandidatPengganti = dataPelanggan.filter(pelanggan => {
        const nik = String(pelanggan.noKTP);
        
        // Lewati jika NIK sudah ada di rencana utama
        if (niksDalamRencana.has(nik)) return false;

        const usageDiPangkalanIni = globalUsageMap.get(nik)?.get(sheetName) || 0;
        const limitPangkalan = (pelanggan.customerTypes === 'Usaha Mikro') ? config.aturanBisnis.batasUsahaMikro : config.aturanBisnis.batasPerPangkalan;
        
        // Validasi jarak hari minimum transaksi
        let lolosJarakHari = false;
        const tglTerakhirDiPangkalanIni = lastTxDateMap.get(nik)?.get(sheetName);

        if (!tglTerakhirDiPangkalanIni) {
            lolosJarakHari = true; // Lolos jika belum pernah transaksi
        } else {
            const today = new Date();
            today.setHours(0, 0, 0, 0);
            const lastTxDate = parseTanggalTransaksi(tglTerakhirDiPangkalanIni);
            if(lastTxDate) {
                lastTxDate.setHours(0, 0, 0, 0);
                const selisihHari = Math.floor((today - lastTxDate) / (1000 * 60 * 60 * 24));
                if (selisihHari >= config.aturanBisnis.jarakHariMinimumTransaksi) {
                    lolosJarakHari = true;
                }
            }
        }
        
        const lolosKuotaCache = !(typeof pelanggan.monthly === 'number' && pelanggan.monthly <= 0);
        
        // Pelanggan lolos jika memenuhi semua syarat
        return usageDiPangkalanIni < limitPangkalan && lolosKuotaCache && lolosJarakHari;
    }).sort((a, b) => { // Mengurutkan untuk memprioritaskan yang paling lama tidak bertransaksi
        const tglA_str = lastTxDateMap.get(String(a.noKTP))?.get(sheetName);
        const tglB_str = lastTxDateMap.get(String(b.noKTP))?.get(sheetName);
        
        const tglA = tglA_str ? parseTanggalTransaksi(tglA_str)?.getTime() || 0 : 0;
        const tglB = tglB_str ? parseTanggalTransaksi(tglB_str)?.getTime() || 0 : 0;

        return tglA - tglB;
    });
    // --- [END] Logika Baru ---

    console.log(chalk.gray(`   > Ditemukan ${kandidatPengganti.length} kandidat pengganti yang memenuhi syarat.`));

    console.log(chalk.blue("\n2. Memulai eksekusi transaksi..."));
    const { summary } = await runTransactionInputProcess(token, profile, dataToInput, kandidatPengganti);
    
    console.log(chalk.bold.green("\n\n🏁 Eksekusi Rencana Selesai."));
    ui.tampilkanTabelRingkasan(summary, "Ringkasan Akhir Eksekusi");
    console.log(`⏱️  Total Waktu Eksekusi: ${ui.formatWaktuProses(Date.now() - startTime)}`);
}

async function runTransactionInputProcess(token, profile, dataDariRencana = null, kandidatPengganti = []) {
    const startTime = Date.now();
    const isModeCerdas = !!dataDariRencana;
    
    if (!isModeCerdas) {
        console.log(chalk.bold.yellow(`\n--- [Utilitas] Input Transaksi Manual ---`));
    }

    const pathMasterPelanggan = config.filePaths.masterPelanggan;
    let dataPelanggan = excel.bacaFile(pathMasterPelanggan) || [];
    const pelangganSet = new Set(dataPelanggan.map(p => String(p.noKTP)));

    let productInfo;
    try {
        productInfo = await api.getProducts(token);
    } catch (error) {
        console.log(chalk.red.bold(`\n❌ Error memuat info produk: ${error.message}`));
        return { summary: {} };
    }

    const sheetName = profile.storeName.replace(/[\\/*?:"\[\]]/g, '').substring(0, 31);
    const { data: historicalData, workbook } = excel.bacaLog(config.filePaths.masterLogTransaksi, sheetName);
    const dataToInput = dataDariRencana || excel.bacaFile(config.filePaths.inputFileTransaksi);

    if (!dataToInput) return { summary: {} };

    const currentMonth = new Date().getMonth();
    const usageMap = new Map();
    historicalData
        .filter(row => row.status?.startsWith('Sukses') && new Date(String(row.tanggal_transaksi).split(',')[0].split('/').reverse().join('-')).getMonth() === currentMonth)
        .forEach(row => usageMap.set(String(row.noKTP), (usageMap.get(String(row.noKTP)) || 0) + (row.quantity || 1)));

    let currentStock = productInfo.stockAvailable;
    const summary = { 'Sukses': 0, 'Gagal (Verifikasi)': 0, 'Gagal (Transaksi)': 0, 'Dilewati (Batas/Stok/Format)': 0, 'Sukses (Pengganti)': 0, 'Gagal (Pengganti)': 0 };
    const progressBar = ui.buatProgressBar(isModeCerdas ? 'Eksekusi Rencana' : 'Input Manual');
    progressBar.start(dataToInput.length, 0);

    const processSingleTransaction = async (user, isReplacement = false) => {
        const nik = String(user.noKTP);
        const quantity = user.quantity || 1;
        
        const resultRow = {
            noKTP: nik, nama: 'Akan diverifikasi', customerTypes: 'Akan diverifikasi', status: 'Belum diproses',
            tanggal_transaksi: new Date().toLocaleString('id-ID', {timeZone: 'Asia/Jakarta'}),
            pangkalan: profile.storeName, quantity,
        };

        try {
            if (!nik || !quantity || !/^\d{16}$/.test(nik)) throw new Error('Format NIK/Quantity Salah');
            if (quantity > currentStock) throw new Error('Stok Tidak Cukup');

            const verificationData = await api.getVerificationData(nik, token);
            if (!verificationData) throw new Error("Verifikasi NIK Gagal");

            resultRow.nama = verificationData.name;
            const customerTypeInfo = verificationData.customerTypes?.[0];
            const custType = customerTypeInfo?.name || 'N/A';
            const sourceTypeIdValue = customerTypeInfo?.sourceTypeId;
            if (!sourceTypeIdValue) throw new Error(`sourceTypeId tidak ditemukan untuk NIK ${nik}`);
            
            resultRow.customerTypes = custType;
            
            const monthlyLimit = (custType === 'Usaha Mikro') ? config.aturanBisnis.batasUsahaMikro : config.aturanBisnis.batasPerPangkalan;
            if ((usageMap.get(nik) || 0) + quantity > monthlyLimit) throw new Error(`Batas ${monthlyLimit} trx/bulan akan terlampaui`);
            
            const payload = {
                products: [{ productId: productInfo.productId, quantity }], token: verificationData.token,
                subsidi: { nik, familyIdEncrypted: verificationData.familyIdEncrypted, category: custType, nama: verificationData.name, channelInject: "tnp2k", sourceTypeId: sourceTypeIdValue },
            };

            const trxData = await api.postTransaction(payload, token);
            if (!trxData.success) throw new Error(trxData.message || "TRANSACTION_INVALID");

            resultRow.status = `Sukses - ID: ${trxData.data.transactionId}`;
            if (isReplacement) summary['Sukses (Pengganti)']++; else summary.Sukses++;
            
            currentStock -= quantity;
            usageMap.set(nik, (usageMap.get(nik) || 0) + quantity);
            
            const pelangganDiMaster = dataPelanggan.find(p => String(p.noKTP) === nik);
            if (pelangganDiMaster) {
                pelangganDiMaster.tanggal_terakhir_transaksi = resultRow.tanggal_transaksi;
            } else if (!pelangganSet.has(nik)) {
                const quota = await api.getQuota(nik, verificationData, token);
                dataPelanggan.push({
                    noKTP: nik, nama: verificationData.name, customerTypes: custType,
                    tanggal_terakhir_transaksi: resultRow.tanggal_transaksi,
                    daily: quota?.daily ?? 0, monthly: quota?.monthly ?? 0,
                    family: quota?.family ?? 'N/A', all: quota?.all ?? 0,
                    terakhir_dicek: new Date().toISOString()
                });
                pelangganSet.add(nik);
            }
            historicalData.push(resultRow);
            return true;

        } catch (error) {
            resultRow.status = `Gagal - ${error.message}`;
            if (isReplacement) {
                summary['Gagal (Pengganti)']++;
            } else {
                if(error.message === "Verifikasi NIK Gagal") summary['Gagal (Verifikasi)']++;
                else if (error.message.includes('Batas') || error.message.includes('Stok') || error.message.includes('Format')) summary['Dilewati (Batas/Stok/Format)']++;
                else summary['Gagal (Transaksi)']++;
            }
            historicalData.push(resultRow);
            return false;
        }
    };

    for (const [index, user] of dataToInput.entries()) {
        progressBar.increment();
        const success = await processSingleTransaction(user, false);
        
        // --- BLOK PERBAIKAN UNTUK KUANTITAS PENGGANTI ---
        if (!success && isModeCerdas) {
    const failedQuantity = user.quantity || 1; 

    if (kandidatPengganti.length > 0) {
        const pengganti = kandidatPengganti.shift();
        const tipePengganti = pengganti.customerTypes || 'Rumah Tangga';

        // 1. Tentukan batas kuantitas MAKSIMAL yang diizinkan sesuai TIPE pelanggan pengganti
        let maksKuantitasPerTransaksi;
        if (tipePengganti === 'Rumah Tangga') {
            // Rumah Tangga punya batas keras per transaksi, kita set misal 2
            maksKuantitasPerTransaksi = 2; 
        } else {
            // Untuk Usaha Mikro, batas per transaksi bisa fleksibel, kita batasi dengan kuantitas yg gagal
            maksKuantitasPerTransaksi = failedQuantity;
        }

        // 2. Tentukan kuantitas final dengan mempertimbangkan SEMUA batasan yang ada
        const kuantitasFinal = Math.min(
            failedQuantity,             // Tidak boleh > dari kuantitas yang gagal
            currentStock,               // Tidak boleh > dari stok
            pengganti.daily || 0,       // Tidak boleh > dari kuota harian pengganti
            pengganti.monthly || 0,     // Tidak boleh > dari kuota bulanan pengganti
            maksKuantitasPerTransaksi   // WAJIB: Tidak boleh > dari batas tipe pelanggannya
        );

        if (kuantitasFinal > 0) {
            console.log(chalk.cyan(`\n   > Transaksi untuk ${user.nama} (qty: ${failedQuantity}) gagal. Mencoba pengganti: ${pengganti.nama} (Tipe: ${tipePengganti}) dengan kuantitas ${kuantitasFinal}`));
            const penggantiDenganKuantitas = { ...pengganti, quantity: kuantitasFinal };
            await processSingleTransaction(penggantiDenganKuantitas, true);
        } else {
             console.log(chalk.red(`\n   > Transaksi gagal. Kandidat pengganti ${pengganti.nama} tidak punya cukup stok/kuota untuk kuantitas ${failedQuantity}.`));
        }

        } else {
                console.log(chalk.red('\n   > Transaksi gagal, tidak ada kandidat pengganti tersisa.'));
            }
        }
        // --- BLOK PERBAIKAN SELESAI ---

        if (currentStock <= 0) {
             progressBar.stop();
             console.log(chalk.red.bold("\n\n🛑 Stok habis, proses dihentikan."));
             break;
        };
        if (index < dataToInput.length - 1) {
            await jeda(config.jeda.transaksi.minDetik, config.jeda.transaksi.maksDetik);
        }
    }
    
    progressBar.stop();
    
    excel.tulisLog(config.filePaths.masterLogTransaksi, workbook, sheetName, historicalData);
    excel.tulisLog(config.filePaths.masterPelanggan, xlsx.utils.book_new(), "Pelanggan", dataPelanggan);
    
    if(!isModeCerdas) {
        console.log(chalk.green.bold("\n\n✅ Proses Selesai. Master Log & Pelanggan telah diperbarui."));
        ui.tampilkanTabelRingkasan(summary, "Ringkasan Input Transaksi");
    }

    return { summary };
}

async function runTambahPelangganSatuan(token) {
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- [Manajemen] Menambahkan Pelanggan Satuan ---"));
    
    let tambahLagi = true;

    while(tambahLagi) {
        try {
            const { nik } = await ui.promptTambahPelangganSatuan();
            const dataPelanggan = excel.bacaFile(config.filePaths.masterPelanggan) || [];
            const pelangganSet = new Set(dataPelanggan.map(p => String(p.noKTP)));
            if (pelangganSet.has(nik)) throw new Error(`Pelanggan dengan NIK ${nik} sudah ada di dalam master.`);
            
            console.log(chalk.blue(`\nMemverifikasi NIK ${nik} ke server...`));
            const verificationData = await api.getVerificationData(nik, token);
            if (!verificationData) throw new Error("Verifikasi NIK Gagal. Pastikan NIK terdaftar di sistem subsidi.");

            const quota = await api.getQuota(nik, verificationData, token);
            const custType = verificationData.customerTypes?.[0]?.name || 'N/A';

            const newPelanggan = {
                noKTP: nik, nama: verificationData.name, customerTypes: custType,
                tanggal_terakhir_transaksi: 'Belum Pernah Transaksi',
                daily: quota?.daily ?? 0, monthly: quota?.monthly ?? 0,
                family: quota?.family ?? 'N/A', all: quota?.all ?? 0,
                terakhir_dicek: new Date().toISOString()
            };

            dataPelanggan.push(newPelanggan);
            excel.tulisLog(config.filePaths.masterPelanggan, xlsx.utils.book_new(), "Pelanggan", dataPelanggan);
            console.log(chalk.green.bold(`\n✅ Pelanggan "${verificationData.name}" berhasil ditambahkan ke master.`));
        } catch (error) {
            console.log(chalk.red.bold(`\n❌ Gagal: ${error.message}`));
        }

        const { continuePrompt } = await inquirer.prompt([{
            type: 'confirm',
            name: 'continuePrompt',
            message: 'Tambah pelanggan satuan lagi?',
            default: true
        }]);
        tambahLagi = continuePrompt;
    }

    console.log(`\n⏱️  Total Waktu Sesi Tambah Pelanggan: ${ui.formatWaktuProses(Date.now() - startTime)}`);
}

async function runCekAndUpdateKuotaPelanggan(token) {
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- Memperbarui Kuota di Master Pelanggan ---"));
    
    const dataPelanggan = excel.bacaFile(config.filePaths.masterPelanggan);
    if (!dataPelanggan) {
        console.log(chalk.red.bold(`\n❌ File master pelanggan tidak ditemukan.`));
        return;
    }

    console.log(chalk.blue(`\nMengambil data kuota terbaru untuk ${dataPelanggan.length} pelanggan...`));
    const progressBar = ui.buatProgressBar('Update Kuota');
    progressBar.start(dataPelanggan.length, 0);

    for (const pelanggan of dataPelanggan) {
        progressBar.increment();
        const nik = String(pelanggan.noKTP);
        try {
            const verificationData = await api.getVerificationData(nik, token);
            if (verificationData) {
                const quota = await api.getQuota(nik, verificationData, token);
                if (quota) {
                    pelanggan.daily = quota.daily;
                    pelanggan.monthly = quota.monthly;
                    pelanggan.family = quota.family;
                    pelanggan.all = quota.all;
                    pelanggan.terakhir_dicek = new Date().toISOString();
                }
            }
        } catch(e) {
            console.log(chalk.red(`\nError saat memproses NIK ${nik}: ${e.message}`));
        }
        await jeda(config.jeda.cekKuota.minDetik, config.jeda.cekKuota.maksDetik);
    }
    progressBar.stop();
    
    excel.tulisLog(config.filePaths.masterPelanggan, xlsx.utils.book_new(), "Pelanggan", dataPelanggan);
    console.log(chalk.green.bold("\n\n✅ Proses Selesai. Data kuota di master pelanggan telah diperbarui."));
    console.log(`⏱️  Waktu Proses: ${ui.formatWaktuProses(Date.now() - startTime)}`);
}

async function runSinkronisasiLaporan(token, profile) {
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- [Laporan] Sinkronisasi Data Transaksi ---"));
    try {
        const { periodePilihan } = await inquirer.prompt([
            {
                type: 'list',
                name: 'periodePilihan',
                message: 'Pilih periode laporan untuk sinkronisasi:',
                choices: [
                    { name: 'Hari Ini', value: 'hariIni' },
                    { name: 'Pilih Rentang Tanggal Manual', value: 'rentang' },
                ],
            },
        ]);

        let startDate, endDate;

        if (periodePilihan === 'hariIni') {
            startDate = new Date();
            endDate = new Date();
            startDate.setHours(0, 0, 0, 0);
            endDate.setHours(23, 59, 59, 999);
            console.log(chalk.blue(`\nMenjalankan sinkronisasi untuk hari ini: ${startDate.toLocaleDateString('id-ID')}`));
        } else {
            const promptedDates = await ui.promptPilihRentangTanggal();
            startDate = promptedDates.startDate;
            endDate = promptedDates.endDate;
        }
        
        const tglMulai = startDate.toISOString().split('T')[0];
        const tglSelesai = endDate.toISOString().split('T')[0];
        
        console.log(chalk.blue("\n1. Membaca data Master Pelanggan lokal..."));
        let dataPelanggan = excel.bacaFile(config.filePaths.masterPelanggan) || [];
        if(dataPelanggan.length === 0) {
            console.log(chalk.yellow("   > Peringatan: File Master Pelanggan kosong atau tidak ditemukan."));
        }
        
        console.log(chalk.blue(`\n2. Mengambil laporan dari server...`));
        const serverCustomers = await api.getTransactionsReport(token, tglMulai, tglSelesai);
        if (serverCustomers.length === 0) {
            console.log(chalk.yellow("Tidak ada transaksi ditemukan di server pada periode ini."));
            return;
        }

        console.log(chalk.blue("\n3. Membaca dan menyinkronkan log transaksi lokal..."));
        const sheetName = profile.storeName.replace(/[\\/*?:"\[\]]/g, '').substring(0, 31);
        const { data: localData, workbook } = excel.bacaLog(config.filePaths.masterLogTransaksi, sheetName);
        const localTransactionIds = new Set(localData.map(row => (row.status || '').match(/ID: ([\w-]+)/)?.[1]).filter(Boolean));

        let transactionsAdded = 0;
        const progressBar = ui.buatProgressBar('Sinkronisasi');
        progressBar.start(serverCustomers.length, 0);

        for (const customer of serverCustomers) {
            const serverTransactions = await api.getTransactionsByCustomer(token, tglMulai, tglSelesai, customer.customerReportId);
            
            for (const tx of serverTransactions) {
                if (!localTransactionIds.has(tx.transactionId)) {
                    const detail = await api.getTransactionDetail(token, tx.transactionId);
                    
                    const maskedNik = tx.nationalityId || detail.subsidi.nik;
                    const customerName = detail.subsidi.nama;
                    const customerCategory = detail.subsidi.category;
                    const cleanName = (name) => name.trim().replace(/\s+/g, ' ').toLowerCase();
                    const serverNameClean = cleanName(customerName);
                    
                    let matchedCustomer = dataPelanggan.find(p => cleanName(p.nama) === serverNameClean && String(p.noKTP).slice(0, 3) === maskedNik.slice(0, 3) && String(p.noKTP).slice(-3) === maskedNik.slice(-3) );
                    
                    let fullNik = maskedNik;
                    if (matchedCustomer) { fullNik = matchedCustomer.noKTP; }
                    
                    if (fullNik !== maskedNik && !fullNik.includes('x')) {
                        console.log(chalk.cyan(`\n   > Menambahkan transaksi hilang: ID ${tx.transactionId} untuk ${customerName} (NIK: ${fullNik})`));
                    } else {
                        console.log(chalk.yellow(`\n   > Menambahkan transaksi hilang: ID ${tx.transactionId} untuk ${customerName} (NIK LENGKAP TIDAK DITEMUKAN)`));
                    }
                    
                    const newLogRow = {
                        noKTP: fullNik,
                        nama: customerName,
                        customerTypes: customerCategory,
                        status: `Sukses - ID: ${detail.transactionId}`,
                        tanggal_transaksi: detail.subHeader.date,
                        pangkalan: detail.receipt.storeName,
                        quantity: detail.products[0].rawValue.quantity
                    };
                    
                    localData.push(newLogRow);
                    transactionsAdded++;
                }
            }
            progressBar.increment();
            await jeda(0, 1);
        }
        progressBar.stop();

        if (transactionsAdded > 0) {
            console.log(chalk.green.bold(`\n\n✅ Sinkronisasi log selesai. ${transactionsAdded} transaksi baru berhasil ditambahkan.`));
            excel.tulisLog(config.filePaths.masterLogTransaksi, workbook, sheetName, localData);
        } else {
            console.log(chalk.green.bold("\n\n✅ Sinkronisasi log selesai. Log lokal Anda sudah sesuai dengan data server."));
        }
        
        console.log(chalk.blue("\n4. Memperbarui 'tanggal_terakhir_transaksi' di Master Pelanggan..."));
        
        const latestTransactionMap = new Map();
        for(const trx of localData){
            if(trx.noKTP && trx.status?.startsWith('Sukses')){
                const nik = String(trx.noKTP);
                const trxDateStr = String(trx.tanggal_transaksi).split(',')[0].split(' ')[0];
                const parts = trxDateStr.split('/');
                let trxDate;
                if(parts.length === 3) { // Format dd/MM/yyyy
                    trxDate = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
                } else { 
                    trxDate = new Date(trx.tanggal_transaksi);
                }

                if(!isNaN(trxDate.getTime())){
                    const existing = latestTransactionMap.get(nik);
                    if(!existing || trxDate > existing.date){
                        latestTransactionMap.set(nik, { date: trxDate, dateString: trx.tanggal_transaksi });
                    }
                }
            }
        }
        
        let masterPelangganUpdated = false;
        for(const pelanggan of dataPelanggan){
            const nik = String(pelanggan.noKTP);
            const latestTrx = latestTransactionMap.get(nik);
            if(latestTrx && pelanggan.tanggal_terakhir_transaksi !== latestTrx.dateString){
                pelanggan.tanggal_terakhir_transaksi = latestTrx.dateString;
                masterPelangganUpdated = true;
            }
        }
        
        if(masterPelangganUpdated){
            excel.tulisLog(config.filePaths.masterPelanggan, xlsx.utils.book_new(), "Pelanggan", dataPelanggan);
            console.log(chalk.green.bold("✅ Master Pelanggan berhasil diperbarui dengan tanggal transaksi terbaru."));
        } else {
            console.log(chalk.green("   > Master Pelanggan sudah sinkron, tidak ada pembaruan tanggal transaksi."));
        }

    } catch (error) {
        console.log(chalk.red.bold(`\n❌ Terjadi kesalahan saat sinkronisasi: ${error.message}`));
    } finally {
        console.log(`\n⏱️  Waktu Proses: ${ui.formatWaktuProses(Date.now() - startTime)}`);
    }
}

async function runValidasiRencana(token, profile) { 
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- [Utilitas] Analisis Kelayakan Rencana (Validasi Lokal) ---"));
    
    const rencana = excel.bacaFile(config.filePaths.rencanaTransaksi);
    if (!rencana) return;

    console.log(chalk.blue("\n1. Membaca data pendukung (Master Pelanggan & Log Transaksi)..."));
    const dataPelanggan = excel.bacaFile(config.filePaths.masterPelanggan) || [];
    const pelangganMasterMap = new Map(dataPelanggan.map(p => [String(p.noKTP), p]));

    const { data: semuaLog } = excel.bacaSemuaSheetLog(config.filePaths.masterLogTransaksi);
    const sheetName = profile.storeName.replace(/[\\/*?:"\[\]]/g, '').substring(0, 31);
    const logPangkalanIni = semuaLog[sheetName] || [];
    
    const usageMap = new Map();
    const lastTxDateMap = new Map();
    const currentMonth = new Date().getMonth();

    logPangkalanIni
        .filter(row => row.status?.startsWith('Sukses'))
        .forEach(row => {
            const nik = String(row.noKTP);
            const trxDate = new Date(String(row.tanggal_transaksi).split(',')[0].split('/').reverse().join('-'));

            if (trxDate.getMonth() === currentMonth) {
                usageMap.set(nik, (usageMap.get(nik) || 0) + (row.quantity || 1));
            }

            if (!lastTxDateMap.has(nik) || trxDate > lastTxDateMap.get(nik)) {
                lastTxDateMap.set(nik, trxDate);
            }
        });

    console.log(chalk.blue(`2. Menganalisis ${rencana.length} pelanggan dalam rencana (mode lokal)...`));
    const validationResults = [];
    
    for (const itemRencana of rencana) {
        const nik = String(itemRencana.noKTP);
        const quantity = itemRencana.quantity || 1;
        let masalah = '-';

        const dataMaster = pelangganMasterMap.get(nik);
        const customerType = dataMaster?.customerTypes || 'Rumah Tangga';

        if (!/^\d{16}$/.test(nik)) {
            masalah = 'Format NIK salah (bukan 16 digit)';
        } else if (!dataMaster) {
            masalah = 'NIK tidak ditemukan di Master Pelanggan';
        } else {
            const limit = (customerType === 'Usaha Mikro') ? config.aturanBisnis.batasUsahaMikro : config.aturanBisnis.batasPerPangkalan;
            if ((usageMap.get(nik) || 0) + quantity > limit) {
                masalah = `Akan melebihi batas bulanan (${usageMap.get(nik) || 0}/${limit})`;
            }

            const tglTerakhir = lastTxDateMap.get(nik);
            if (tglTerakhir) {
                const today = new Date(); today.setHours(0, 0, 0, 0);
                const selisihHari = Math.floor((today - tglTerakhir) / (1000 * 60 * 60 * 24));
                if (selisihHari < config.aturanBisnis.jarakHariMinimumTransaksi) {
                    masalah = `Baru bertransaksi ${selisihHari} hari yang lalu`;
                }
            }
        }
        
        const status = (masalah !== '-') ? { text: `⚠️ Berpotensi Gagal`, color: 'yellow' } : { text: '✅ Lolos Lokal', color: 'green' };

        validationResults.push({
            nik,
            nama: dataMaster?.nama || itemRencana.nama || 'N/A',
            quantity,
            status: chalk[status.color](status.text),
            masalah
        });
    }
    
    console.log(chalk.bold("\n--- Hasil Analisis Kelayakan Rencana (Lokal) ---"));
    const table = new Table({ head: [chalk.cyan('NIK'), chalk.cyan('Nama'), chalk.cyan('Qty'), chalk.cyan('Status'), chalk.cyan('Keterangan')] });
    validationResults.forEach(res => table.push([res.nik, res.nama, res.quantity, res.status, res.masalah]));
    console.log(table.toString());
    console.log(`\n⏱️  Waktu Proses: ${ui.formatWaktuProses(Date.now() - startTime)}`);
}

async function runManajemenTemplate() {
    console.log(chalk.bold.yellow("\n--- [Utilitas] Membuat File Template Input ---"));
    const ensureDirExists = (filePath) => {
        const dir = filePath.substring(0, filePath.lastIndexOf('/'));
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    }

    const pathRencana = config.filePaths.rencanaTransaksi;
    ensureDirExists(pathRencana);
    if (!fs.existsSync(pathRencana)) {
        excel.tulisLog(pathRencana, xlsx.utils.book_new(), "Rencana", [{ noKTP: '1234567890123456', nama: 'NAMA (opsional)', customerTypes: 'Rumah Tangga (opsional)', quantity: 1, }]);
        console.log(chalk.green(`   ✅ Template "${pathRencana}" berhasil dibuat.`));
    }

    const pathManual = config.filePaths.inputFileTransaksi;
    ensureDirExists(pathManual);
    if (!fs.existsSync(pathManual)) {
        excel.tulisLog(pathManual, xlsx.utils.book_new(), "Sheet1", [{ noKTP: '1234567890123456', quantity: 1 }]);
        console.log(chalk.green(`   ✅ Template "${pathManual}" berhasil dibuat.`));
    }

    const pathPelangganBaru = config.filePaths.inputFilePelangganBaru;
    ensureDirExists(pathPelangganBaru);
     if (!fs.existsSync(pathPelangganBaru)) {
        excel.tulisLog(pathPelangganBaru, xlsx.utils.book_new(), "Sheet1", [{ noKTP: '1234567890123456' }]);
        console.log(chalk.green(`   ✅ Template "${pathPelangganBaru}" berhasil dibuat.`));
    }
}

async function runDiagnosisData(token) {
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- [Diagnosis] Cek & Perbaiki Data Pelanggan ---"));
    try {
        const { startDate, endDate } = await ui.promptPilihRentangTanggal();
        const tglMulai = startDate.toISOString().split('T')[0];
        const tglSelesai = endDate.toISOString().split('T')[0];
        
        console.log(chalk.blue("\n1. Membaca data Master Pelanggan lokal..."));
        const dataPelanggan = excel.bacaFile(config.filePaths.masterPelanggan) || [];
        
        console.log(chalk.blue(`\n2. Mengambil laporan dari server untuk perbandingan...`));
        const serverCustomers = await api.getTransactionsReport(token, tglMulai, tglSelesai);
        if (serverCustomers.length === 0) {
            console.log(chalk.yellow("Tidak ada data di server untuk dibandingkan pada periode ini."));
            return;
        }

        const discrepancies = [];
        const cleanName = (name) => name.trim().replace(/\s+/g, ' ').toLowerCase();

        for (const customer of serverCustomers) {
            const serverName = customer.name;
            const maskedNik = customer.nationalityId;
            const serverNameClean = cleanName(serverName);

            const perfectMatch = dataPelanggan.find(p => cleanName(p.nama) === serverNameClean && String(p.noKTP).slice(0, 3) === maskedNik.slice(0, 3) && String(p.noKTP).slice(-3) === maskedNik.slice(-3) );
            if (!perfectMatch) {
                const nikPatternMatches = dataPelanggan.filter(p => String(p.noKTP).slice(0, 3) === maskedNik.slice(0, 3) && String(p.noKTP).slice(-3) === maskedNik.slice(-3) );
                let suggestion = 'Tidak ada kandidat NIK yang cocok.';
                if (nikPatternMatches.length > 0) {
                    suggestion = nikPatternMatches.map(p => `${p.nama} | ${p.noKTP}`).join(', ');
                }
                discrepancies.push({ serverName, maskedNik, suggestion });
            }
        }

        if (discrepancies.length > 0) {
            console.log(chalk.red.bold(`\n\n❌ Ditemukan ${discrepancies.length} potensi data yang tidak cocok!`));
            const table = new Table({ head: [chalk.cyan('Nama di Server'), chalk.cyan('Pola NIK Server'), chalk.cyan('Saran Perbaikan di Master Lokal Anda')] });
            discrepancies.forEach(d => table.push([d.serverName, d.maskedNik, d.suggestion]));
            console.log(table.toString());
            console.log(chalk.yellow.bold("\nINSTRUKSI:\n1. Buka file MASTER_PELANGGAN.xlsx Anda.\n2. Cari pelanggan berdasarkan NIK pada kolom 'Saran Perbaikan'.\n3. Ubah nama pelanggan di master Anda agar SAMA PERSIS dengan 'Nama di Server'.\n4. Setelah semua diperbaiki, jalankan kembali 'Menu 10: Sinkronisasi'."));
        } else {
            console.log(chalk.green.bold("\n\n✅ Selamat! Semua data pelanggan yang bertransaksi sudah cocok dengan master lokal Anda."));
        }
    } catch (error) {
        console.log(chalk.red.bold(`\n❌ Terjadi kesalahan saat diagnosis: ${error.message}`));
    } finally {
        console.log(`⏱️  Waktu Proses: ${ui.formatWaktuProses(Date.now() - startTime)}`);
    }
}

async function runSearchLocalMaster() {
    const startTime = Date.now();
    console.log(chalk.bold.yellow("\n--- [Diagnosis] Cari Pelanggan di Master Lokal ---"));

    console.log(chalk.blue("Membaca data Master Pelanggan & Log Transaksi..."));
    const dataPelanggan = excel.bacaFile(config.filePaths.masterPelanggan) || [];
    if (dataPelanggan.length === 0) {
        console.log(chalk.yellow("File Master Pelanggan kosong atau tidak ditemukan."));
        return;
    }

    const { data: semuaLog } = excel.bacaSemuaSheetLog(config.filePaths.masterLogTransaksi);
    const monthlyUsageMap = new Map();
    const currentMonth = new Date().getMonth();

    // --- BLOK PERBAIKAN LOGIKA TANGGAL ---
    for (const sheet in semuaLog) {
        semuaLog[sheet]
            .filter(row => row.status?.startsWith('Sukses'))
            .forEach(row => {
                // Logika parsing tanggal yang lebih kuat
                const dateString = String(row.tanggal_transaksi).split(' ')[0].replace(',', '');
                const dateParts = dateString.split('/');
                if (dateParts.length === 3) {
                    const trxDate = new Date(`${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`);
                    if (!isNaN(trxDate.getTime()) && trxDate.getMonth() === currentMonth) {
                        const nik = String(row.noKTP);
                        monthlyUsageMap.set(nik, (monthlyUsageMap.get(nik) || 0) + (row.quantity || 1));
                    }
                }
            });
    }
    // --- AKHIR BLOK PERBAIKAN ---

    let cariLagi = true;
    while (cariLagi) {
        try {
            const { nama } = await ui.promptSearchByName();
            const searchInput = nama.trim().toLowerCase();
            const results = dataPelanggan.filter(p => p.nama.trim().toLowerCase().includes(searchInput));
            
            if (results.length > 0) {
                console.log(chalk.green(`\n✅ Ditemukan ${results.length} hasil untuk pencarian "${nama}":`));
                
                const table = new Table({ 
                    head: [chalk.cyan('Nama Lengkap'), chalk.cyan('NIK'), chalk.cyan('Terakhir Transaksi'), chalk.cyan('Qty Bulan Ini')] 
                });

                results.forEach(p => {
                    const nik = String(p.noKTP);
                    const qtyBulanIni = monthlyUsageMap.get(nik) || 0;
                    let tglTerakhir = p.tanggal_terakhir_transaksi || '-';
                    // Membersihkan tampilan tanggal terakhir
                    tglTerakhir = tglTerakhir.split(' ')[0].replace(',', '');

                    table.push([p.nama, nik, tglTerakhir, qtyBulanIni]);
                });
                console.log(table.toString());

            } else {
                console.log(chalk.red(`\n❌ Tidak ada pelanggan yang cocok dengan pencarian "${nama}" di dalam file master Anda.`));
            }
        } catch (error) {
            console.log(chalk.red.bold(`\n❌ Terjadi kesalahan saat mencari: ${error.message}`));
        }

        const { continuePrompt } = await inquirer.prompt([{
            type: 'confirm',
            name: 'continuePrompt',
            message: 'Cari pelanggan lain di master lokal?',
            default: true
        }]);
        cariLagi = continuePrompt;
    }
    console.log(`\n⏱️  Total Waktu Sesi Pencarian: ${ui.formatWaktuProses(Date.now() - startTime)}`);
}

async function startSystem() {
    await ui.initializeDependencies();
    chalk = (await import('chalk')).default;
    inquirer = (await import('inquirer')).default;
    ui.tampilkanHeader();
    let token = await ui.dapatkanToken();
    if (!token) return;

    let profile;
    while(true) {
        console.log(chalk.blue("\nMemverifikasi token & mengambil profil..."));
        profile = await api.getProfileInfo(token);
        if (profile) { console.log(chalk.green("Token valid.")); break; }
        
        console.log(chalk.red.bold("Token tidak valid atau gagal."));
        if(fs.existsSync(config.filePaths.tokenCache)) fs.unlinkSync(config.filePaths.tokenCache);
        const { retry } = await inquirer.prompt([{ type: 'confirm', name: 'retry', message: 'Coba lagi?', default: true }]);
        if (retry) { token = await ui.dapatkanToken(); } else { return; }
    }
    
    while (true) {
        let productInfo;
        try {
            productInfo = await api.getProducts(token);
        } catch(e) {
            console.log(chalk.red.bold(`\n Gagal memuat info produk/stok. Error: ${e.message}`));
            console.log(chalk.yellow("Silakan coba ganti token atau mulai ulang."));
            productInfo = null;
        }

        ui.tampilkanHeader();
        if(profile && productInfo) ui.tampilkanDashboard(profile, productInfo);
        else console.log(chalk.red("\nTidak dapat menampilkan dashboard, info pangkalan/produk gagal dimuat."));

        const { menuChoice } = await ui.tampilkanMenuUtama();

        if (['3', '4', '5', '10'].includes(menuChoice)) {
            await backupMasterLog();
        }

        switch (menuChoice) {
            case '1': await buatRencanaTransaksi(token, profile); break;
            case '2': await runValidasiRencana(token, profile); break;
            case '3': await eksekusiRencanaTransaksi(token, profile); break;
            case '4': await runTransaksiLangsung(token, profile); break;
            case '5': await runTransactionInputProcess(token, profile); break;
            case '6': await runTambahPelangganBaru(token); break;
            case '7': await runTambahPelangganSatuan(token); break;
            case '8': await runBuatMasterPelanggan(); break;
            case '9': await runCekAndUpdateKuotaPelanggan(token); break;
            case '10': await runSinkronisasiLaporan(token, profile); break;
            case '11': await runManajemenTemplate(); break;
            case '12': await runDiagnosisData(token); break;
            case '13': await runSearchLocalMaster(); break;
            case '14': await startSystem(); return;
            case '15': 
                console.log(chalk.yellow("\nTerima kasih! Sampai jumpa."));
                process.exit(0);
        }

        const { returnToMenu } = await inquirer.prompt([{ type: 'confirm', name: 'returnToMenu', message: 'Kembali ke menu utama?', default: true, }]);
        if (!returnToMenu) {
            console.log(chalk.yellow("\nTerima kasih! Sampai jumpa."));
            process.exit(0);
        }
    }
}

startSystem();