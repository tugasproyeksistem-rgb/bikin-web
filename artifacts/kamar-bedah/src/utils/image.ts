/* src/utils/image.ts
   Fungsi utilitas murni (Canvas/File API, tanpa React) untuk kompresi &
   rotasi gambar sebelum dikirim ke endpoint OCR. Dipindah dari App.tsx
   (langkah 3 refactor — pisah utils, batch 5), TIDAK ADA perubahan logika. */

/** Kompresi + resize gambar di client sebelum dikirim ke Worker.
 *  Mengecilkan payload — penting untuk jaringan RS yang lambat, dan
 *  mengurangi token/neuron yang dipakai model vision (lebih hemat
 *  kuota harian). */
export function compressImageToBase64(
  file: File,
  maxDim = 1400,
  quality = 0.8
): Promise<{ base64: string; mediaType: string }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Gagal membaca file"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("Gagal memuat gambar"));
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          const scale = maxDim / Math.max(width, height);
          width = Math.round(width * scale);
          height = Math.round(height * scale);
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        if (!ctx) return reject(new Error("Canvas tidak didukung"));
        ctx.drawImage(img, 0, 0, width, height);
        const dataUrl = canvas.toDataURL("image/jpeg", quality);
        const base64 = dataUrl.split(",")[1] || "";
        resolve({ base64, mediaType: "image/jpeg" });
      };
      img.src = reader.result as string;
    };
    reader.readAsDataURL(file);
  });
}

/** SARAN PENINGKATAN #7: merotasi gambar 90°/180°/270° via canvas SEBELUM
 *  dikirim ke OCR. Dipakai di langkah preview ScanFormulir supaya foto yang
 *  miring/terbalik (kasus umum saat foto formulir kertas buru-buru) bisa
 *  diluruskan dulu tanpa perlu foto ulang dari awal — akurasi OCR untuk
 *  teks yang terbalik/miring 90° biasanya jauh lebih buruk. */
export function rotateImageFile(file: File, degrees: number): Promise<File> {
  if(((degrees % 360) + 360) % 360 === 0) return Promise.resolve(file);
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Gagal membaca file untuk rotasi"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("Gagal memuat gambar untuk rotasi"));
      img.onload = () => {
        const swap = degrees % 180 !== 0;
        const canvas = document.createElement("canvas");
        canvas.width  = swap ? img.height : img.width;
        canvas.height = swap ? img.width  : img.height;
        const ctx = canvas.getContext("2d");
        if(!ctx) return reject(new Error("Canvas tidak didukung"));
        ctx.translate(canvas.width/2, canvas.height/2);
        ctx.rotate((degrees * Math.PI) / 180);
        ctx.drawImage(img, -img.width/2, -img.height/2);
        canvas.toBlob((blob) => {
          if(!blob) return reject(new Error("Gagal membuat gambar hasil rotasi"));
          resolve(new File([blob], file.name || "formulir.jpg", {type:"image/jpeg"}));
        }, "image/jpeg", 0.92);
      };
      img.src = reader.result as string;
    };
    reader.readAsDataURL(file);
  });
}
