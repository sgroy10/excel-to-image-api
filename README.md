# Excel to Image API

Microservice to convert Excel BOM sheets to PNG images with full formatting preserved (images, colors, merged cells, borders, etc.)

## API Endpoints

### `GET /`
Health check endpoint.

### `POST /convert`
Convert Excel file to PNG image (single page).

**Parameters:**
- `file` (required): Excel file (.xlsx, .xls, .xlsm)
- `dpi` (optional): Image resolution (default: 200)
- `page` (optional): Page number to convert (default: 1)

**Returns:** PNG image

### `POST /convert-all`
Convert all pages of Excel file to base64 images.

**Parameters:**
- `file` (required): Excel file
- `dpi` (optional): Image resolution (default: 200)

**Returns:** JSON with base64 encoded images

---

## Deploy to Railway (Recommended)

### Step 1: Create GitHub Repository

1. Go to [github.com/new](https://github.com/new)
2. Create a new repository named `excel-to-image-api`
3. Upload these 4 files:
   - `main.py`
   - `requirements.txt`
   - `Dockerfile`
   - `README.md`

### Step 2: Deploy on Railway

1. Go to [railway.app](https://railway.app) and sign in with GitHub
2. Click **"New Project"** → **"Deploy from GitHub repo"**
3. Select your `excel-to-image-api` repository
4. Railway will auto-detect the Dockerfile and start building
5. Once deployed, go to **Settings** → **Networking** → **Generate Domain**
6. You'll get a URL like: `https://excel-to-image-api-production.up.railway.app`

### Step 3: Test Your API

```bash
# Health check
curl https://YOUR-RAILWAY-URL.up.railway.app/

# Convert Excel to image
curl -X POST "https://YOUR-RAILWAY-URL.up.railway.app/convert" \
  -F "file=@DLER00510.xlsx" \
  -F "dpi=200" \
  --output result.png
```

---

## Usage in Lovable (Frontend Code)

### Basic Usage - Get Image Blob

```typescript
const convertExcelToImage = async (file: File): Promise<Blob> => {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('dpi', '200');
  
  const response = await fetch('https://YOUR-RAILWAY-URL.up.railway.app/convert', {
    method: 'POST',
    body: formData,
  });
  
  if (!response.ok) {
    throw new Error('Conversion failed');
  }
  
  return await response.blob();
};

// Usage
const imageBlob = await convertExcelToImage(excelFile);
const imageUrl = URL.createObjectURL(imageBlob);
```

### Display in React Component

```tsx
import { useState } from 'react';

const API_URL = 'https://YOUR-RAILWAY-URL.up.railway.app';

export function BOMConverter() {
  const [imageUrl, setImageUrl] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    try {
      const formData = new FormData();
      formData.append('file', file);
      formData.append('dpi', '200');

      const response = await fetch(`${API_URL}/convert`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) throw new Error('Conversion failed');

      const blob = await response.blob();
      setImageUrl(URL.createObjectURL(blob));
    } catch (error) {
      console.error('Error:', error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div>
      <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
      {loading && <p>Converting...</p>}
      {imageUrl && <img src={imageUrl} alt="BOM Sheet" />}
    </div>
  );
}
```

### Download Image

```typescript
const downloadImage = async (file: File, filename: string) => {
  const blob = await convertExcelToImage(file);
  const url = URL.createObjectURL(blob);
  
  const a = document.createElement('a');
  a.href = url;
  a.download = filename.replace(/\.xlsx?$/, '.png');
  a.click();
  
  URL.revokeObjectURL(url);
};
```

---

## Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Make sure LibreOffice is installed
# Ubuntu/Debian: sudo apt install libreoffice poppler-utils
# Mac: brew install libreoffice poppler

# Run server
uvicorn main:app --reload --port 8000

# Test
curl http://localhost:8000/
```

---

## Pricing

**Railway Free Tier:**
- 500 hours/month execution time
- Enough for ~1000-2000 conversions/month

**Railway Pro ($5/month):**
- Unlimited execution
- Better performance

---

## Support

Built for JewelRender.in BOM sheet conversion.
