import React, { useState, useEffect, useRef } from 'react';
import JsBarcode from 'jsbarcode';

function BarcodeModal({ product, isOpen, onClose }) {
  const [selectedSize, setSelectedSize] = useState('large'); // 'large' (3.8x2.5cm) හෝ 'small' (3.0x2.0cm)
  const [quantity, setQuantity] = useState(1);
  const barcodeRefs = useRef([]);

  if (!isOpen || !product) return null;

  // තෝරාගත් ප්‍රමාණය අනුව ලේබල් එකක සයිස් එක CSS සඳහා ලබාදීම
  const dimensions = selectedSize === 'large' 
    ? { width: '3.8cm', height: '2.5cm', fontSize: '11px', barcodeHeight: 28 }
    : { width: '3.0cm', height: '2.0cm', fontSize: '9px', barcodeHeight: 20 };

  // ලැබෙන ප්‍රමාණයට අදාළව බාර්කෝඩ් කැන්වස් (Canvas) යාවත්කාලීන කිරීම
  useEffect(() => {
    if (isOpen) {
      for (let i = 0; i < quantity; i++) {
        if (barcodeRefs.current[i]) {
          try {
            JsBarcode(barcodeRefs.current[i], product.barcode || product.code || '123456', {
              format: "CODE128",
              lineColor: "#000",
              width: selectedSize === 'large' ? 1.4 : 1.1,
              height: dimensions.barcodeHeight,
              displayValue: false
            });
          } catch (e) {
            console.error(e);
          }
        }
      }
    }
  }, [quantity, selectedSize, isOpen, product]);

  const handlePrint = () => {
    const printContent = document.getElementById('barcode-print-area').innerHTML;
    
    const printWindow = window.open('', '_blank', 'width=800,height=600');
    printWindow.document.write(`
      <html>
        <head>
          <title>Print Barcodes</title>
          <style>
            @page { size: A4; margin: 10mm; }
            body { font-family: Arial, sans-serif; margin: 0; padding: 0; }
            .grid-container {
              display: grid;
              grid-template-columns: ${selectedSize === 'large' ? 'repeat(5, 3.8cm)' : 'repeat(6, 3.0cm)'};
              gap: 2mm;
              justify-content: start;
            }
            .barcode-cell {
              width: ${dimensions.width};
              height: ${dimensions.height};
              border: 1px dashed #999; /* කපන ඉරි (Cutting Lines) */
              box-sizing: border-box;
              padding: 3px 2px;
              text-align: center;
              display: flex;
              flex-direction: column;
              align-items: center;
              justify-content: space-between;
              overflow: hidden;
              background: #fff;
            }
            .barcode-code { font-size: ${selectedSize === 'large' ? '10px' : '8px'}; font-weight: bold; margin: 0; line-height: 1; }
            .item-name { font-size: ${selectedSize === 'large' ? '9px' : '7px'}; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; width: 100%; margin: 0; line-height: 1; }
            .barcode-img-container { margin: 1px 0; display: flex; align-items: center; justify-content: center; }
            img { max-width: 100%; height: auto; max-height: ${dimensions.barcodeHeight}px; display: block; }
            .footer-info { display: flex; justify-content: space-between; width: 100%; padding: 0 4px; box-sizing: border-box; align-items: center; line-height: 1; }
            .item-price { font-size: ${selectedSize === 'large' ? '10px' : '8px'}; font-weight: bold; margin: 0; }
            .shop-name { font-size: ${selectedSize === 'large' ? '8px' : '6px'}; margin: 0; }
          </style>
        </head>
        <body>
          <div class="grid-container">
            ${printContent}
          </div>
          <script>
            window.onload = function() {
              window.print();
              window.close();
            }
          </script>
        </body>
      </html>
    `);
    printWindow.document.close();
  };

  return (
    <div style={{ position: 'fixed', top: 0, left: 0, width: '100%', height: '100%', background: 'rgba(0,0,0,0.5)', display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 1000 }}>
      <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', width: '500px', maxWidth: '90%' }}>
        <h3 style={{ margin: '0 0 15px 0', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <span>Barcode මුද්‍රණය — {product.name}</span>
          <button onClick={onClose} style={{ background: 'none', border: 'none', fontSize: '18px', cursor: 'pointer' }}>&times;</button>
        </h3>

        {/* සයිස් තෝරන කොටස */}
        <div style={{ marginBottom: '15px' }}>
          <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold', fontSize: '14px' }}>Barcode Label Size</label>
          <div style={{ display: 'flex', gap: '10px' }}>
            <div 
              onClick={() => setSelectedSize('large')}
              style={{ flex: 1, border: selectedSize === 'large' ? '2px solid #007bff' : '1px solid #ccc', padding: '10px', borderRadius: '6px', cursor: 'pointer', background: selectedSize === 'large' ? '#f0f7ff' : '#fff' }}
            >
              <div style={{ fontWeight: 'bold', fontSize: '13px' }}>(01) විශාල</div>
              <div style={{ fontSize: '11px', color: '#666' }}>පළල: 3.8cm | උස: 2.5cm</div>
              <div style={{ fontSize: '10px', color: '#007bff' }}>A4 → 5 × 11 = 55 ලේබල්</div>
            </div>

            <div 
              onClick={() => setSelectedSize('small')}
              style={{ flex: 1, border: selectedSize === 'small' ? '2px solid #007bff' : '1px solid #ccc', padding: '10px', borderRadius: '6px', cursor: 'pointer', background: selectedSize === 'small' ? '#f0f7ff' : '#fff' }}
            >
              <div style={{ fontWeight: 'bold', fontSize: '13px' }}>(02) කුඩා</div>
              <div style={{ fontSize: '11px', color: '#666' }}>පළල: 3.0cm | උස: 2.0cm</div>
              <div style={{ fontSize: '10px', color: '#007bff' }}>A4 → 6 × 14 = 84 ලේබල්</div>
            </div>
          </div>
        </div>

        {/* ප්‍රමාණය ඇතුළත් කරන කොටස */}
        <div style={{ marginBottom: '20px' }}>
          <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold', fontSize: '14px' }}>මුද්‍රණය කළ යුතු ප්‍රමාණය (ලේබල් ගණන)</label>
          <input 
            type="number" 
            min="1" 
            max="200" 
            value={quantity} 
            onChange={(e) => setQuantity(Math.max(1, parseInt(e.target.value) || 1))}
            style={{ width: '100%', padding: '8px', boxSizing: 'border-box', borderRadius: '4px', border: '1px solid #ccc' }}
          />
        </div>

        {/* බටන්ස් */}
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
          <button onClick={onClose} style={{ padding: '8px 15px', background: '#e0e0e0', border: 'none', borderRadius: '4px', cursor: 'pointer' }}>අවලංගුයි</button>
          <button onClick={handlePrint} style={{ padding: '8px 20px', background: '#007bff', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontWeight: 'bold' }}>Print PDF</button>
        </div>

        {/* සැඟවුණු ප්‍රින්ට් සැකිල්ල (Hidden Printable Grid Templates) */}
        <div style={{ display: 'none' }}>
          <div id="barcode-print-area">
            {Array.from({ length: quantity }).map((_, index) => (
              <div key={index} className="barcode-cell">
                <div className="barcode-code">{product.barcode || product.code}</div>
                <div className="item-name">{product.name}</div>
                <div className="barcode-img-container">
                  <img ref={el => barcodeRefs.current[index] = el} alt="barcode" />
                </div>
                <div className="footer-info">
                  <div className="item-price">
                    Rs. {product.price || product.selling_price} {product.unit ? `(${product.unit})` : ''}
                  </div>
                  <div className="shop-name">POSshitha Stores</div>
                </div>
              </div>
            ))}
          </div>
        </div>

      </div>
    </div>
  );
}

export default BarcodeModal;