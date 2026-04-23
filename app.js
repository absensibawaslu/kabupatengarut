// GANTI URL INI DENGAN URL WEB APP GOOGLE APPS SCRIPT ANDA
const GAS_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzMeqQv7B2GY9OOLSU18MjjAAXj4VPeVizPspcvIVWlgw7YWAr9EXCrr55V4zxLWBJapg/exec';

// DOM Elements
const digitalClock = document.getElementById('digitalClock');
const currentDate = document.getElementById('currentDate');
const cameraStream = document.getElementById('cameraStream');
const canvas = document.getElementById('canvas');
const photoPreview = document.getElementById('photoPreview');
const retakeBtn = document.getElementById('retakeBtn');
const locationStatus = document.getElementById('locationStatus');
const btnIn = document.getElementById('btnIn');
const btnOut = document.getElementById('btnOut');
const employeeNameInput = document.getElementById('employeeName');
const statusMessage = document.getElementById('statusMessage');

// State
let currentLocation = null;
let currentAddress = "";
let currentPhotoBase64 = null;
let stream = null;

// Initialize Clock
function updateClock() {
    const now = new Date();
    digitalClock.textContent = now.toLocaleTimeString('id-ID', { hour12: false });
    currentDate.textContent = now.toLocaleDateString('id-ID', {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
}
setInterval(updateClock, 1000);
updateClock();

// Initialize Camera
async function startCamera() {
    try {
        stream = await navigator.mediaDevices.getUserMedia({
            video: { facingMode: 'user' },
            audio: false
        });
        cameraStream.srcObject = stream;
        cameraStream.style.display = 'block';
        photoPreview.style.display = 'none';
        retakeBtn.classList.add('hidden');
        currentPhotoBase64 = null;
    } catch (err) {
        console.error("Error accessing camera:", err);
        showStatus('Gagal mengakses kamera. Pastikan izin diberikan.', 'error');
    }
}

function takePhoto() {
    if (!stream) return false;

    canvas.width = cameraStream.videoWidth;
    canvas.height = cameraStream.videoHeight;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(cameraStream, 0, 0, canvas.width, canvas.height);
    currentPhotoBase64 = canvas.toDataURL('image/jpeg', 0.8);

    photoPreview.src = currentPhotoBase64;
    cameraStream.style.display = 'none';
    photoPreview.style.display = 'block';
    retakeBtn.classList.remove('hidden');

    return true;
}

/**
 * Mendapatkan Nama Lokasi (Reverse Geocoding)
 */
async function getAddress(lat, lng) {
    try {
        const response = await fetch(`https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lng}&zoom=18&addressdetails=1`);
        const data = await response.json();
        return data.display_name || "Alamat tidak ditemukan";
    } catch (err) {
        return "Gagal mendapatkan nama alamat";
    }
}

// Initialize Location
function getLocation() {
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(
            async (position) => {
                const lat = position.coords.latitude;
                const lng = position.coords.longitude;
                currentLocation = { lat, lng };

                locationStatus.innerHTML = `<span><i data-lucide="loader"></i> Mencari alamat...</span>`;
                lucide.createIcons();

                currentAddress = await getAddress(lat, lng);

                locationStatus.innerHTML = `<i data-lucide="check-circle"></i> <span>Lokasi: ${currentAddress}</span>`;
                locationStatus.className = 'location-status success';
                lucide.createIcons();
                checkFormValidity();
            },
            (error) => {
                locationStatus.innerHTML = `<i data-lucide="x-circle"></i> <span>Gagal mendapatkan lokasi. Pastikan GPS aktif.</span>`;
                locationStatus.className = 'location-status error';
                lucide.createIcons();
            },
            { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
        );
    } else {
        locationStatus.innerHTML = `<i data-lucide="x-circle"></i> <span>Browser tidak mendukung Geolocation</span>`;
        locationStatus.className = 'location-status error';
        lucide.createIcons();
    }
}

function checkFormValidity() {
    if (employeeNameInput.value.trim() !== '' && currentLocation !== null) {
        btnIn.disabled = false;
        btnOut.disabled = false;
    } else {
        btnIn.disabled = true;
        btnOut.disabled = true;
    }
}

employeeNameInput.addEventListener('input', checkFormValidity);
retakeBtn.addEventListener('click', startCamera);

// Submit Data
async function submitAttendance(type) {
    const name = employeeNameInput.value.trim();
    if (!name) {
        showStatus('Masukkan nama Anda!', 'error');
        return;
    }

    if (!currentPhotoBase64) {
        if (!takePhoto()) {
            showStatus('Gagal mengambil foto.', 'error');
            return;
        }
    }

    if (!currentLocation) {
        showStatus('Lokasi GPS belum tersedia.', 'error');
        return;
    }

    const payload = {
        type: type,
        name: name,
        timestamp: new Date().toISOString(),
        lat: currentLocation.lat,
        lng: currentLocation.lng,
        address: currentAddress,
        photo: currentPhotoBase64
    };

    try {
        btnIn.disabled = true;
        btnOut.disabled = true;
        showStatus('Mengirim data...', 'loading');

        // Gunakan mode: no-cors untuk kestabilan pengiriman dari file lokal
        await fetch(GAS_WEB_APP_URL, {
            method: 'POST',
            mode: 'no-cors',
            body: JSON.stringify(payload),
            headers: { 'Content-Type': 'text/plain;charset=utf-8' }
        });

        showStatus('Data Absen ' + type + ' Terkirim! Cek Spreadsheet Anda.', 'success');

        setTimeout(() => {
            startCamera();
            statusMessage.className = 'status-message';
            checkFormValidity();
        }, 3000);

    } catch (error) {
        console.error("Submit Error:", error);
        showStatus('Gagal mengirim data. Periksa koneksi.', 'error');
        checkFormValidity();
    }
}

function showStatus(message, type) {
    statusMessage.textContent = message;
    statusMessage.className = `status-message show ${type}`;
}

btnIn.addEventListener('click', () => submitAttendance('Masuk'));
btnOut.addEventListener('click', () => submitAttendance('Pulang'));

// On Load
startCamera();
getLocation();
