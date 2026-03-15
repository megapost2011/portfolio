#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
画像鮮鋭化・超解像アプリケーション
ぼやけた画像を拡大してハッキリクッキリに！
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk, ImageFilter, ImageEnhance
import numpy as np
from scipy import ndimage
import os

class ImageSharpenerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("画像鮮鋭化・超解像ツール")
        self.root.geometry("900x700")
        
        self.original_image = None
        self.enhanced_image = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """UIのセットアップ"""
        # メインフレーム
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=0, column=0, columnspan=2, pady=5, sticky=tk.W)
        
        ttk.Button(button_frame, text="📂 画像を開く", command=self.load_image).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="⚡ 鮮鋭化実行", command=self.sharpen_image).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="💾 保存", command=self.save_image).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="🔄 リセット", command=self.reset_image).pack(side=tk.LEFT, padx=5)
        
        # パラメータフレーム
        param_frame = ttk.LabelFrame(main_frame, text="鮮鋭化パラメータ", padding="10")
        param_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        
        # 拡大率
        ttk.Label(param_frame, text="拡大率:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.scale_var = tk.DoubleVar(value=2.0)
        scale_slider = ttk.Scale(param_frame, from_=1.5, to=4.0, variable=self.scale_var, 
                                orient=tk.HORIZONTAL, length=200)
        scale_slider.grid(row=0, column=1, padx=5)
        self.scale_label = ttk.Label(param_frame, text="2.0x")
        self.scale_label.grid(row=0, column=2, padx=5)
        scale_slider.configure(command=lambda v: self.scale_label.configure(text=f"{float(v):.1f}x"))
        
        # 鮮鋭化強度
        ttk.Label(param_frame, text="鮮鋭化強度:").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.sharp_var = tk.DoubleVar(value=3.0)
        sharp_slider = ttk.Scale(param_frame, from_=1.0, to=5.0, variable=self.sharp_var,
                                orient=tk.HORIZONTAL, length=200)
        sharp_slider.grid(row=1, column=1, padx=5)
        self.sharp_label = ttk.Label(param_frame, text="3.0")
        self.sharp_label.grid(row=1, column=2, padx=5)
        sharp_slider.configure(command=lambda v: self.sharp_label.configure(text=f"{float(v):.1f}"))
        
        # エッジ強調
        ttk.Label(param_frame, text="エッジ強調:").grid(row=2, column=0, sticky=tk.W, padx=5)
        self.edge_var = tk.DoubleVar(value=2.0)
        edge_slider = ttk.Scale(param_frame, from_=0.0, to=4.0, variable=self.edge_var,
                               orient=tk.HORIZONTAL, length=200)
        edge_slider.grid(row=2, column=1, padx=5)
        self.edge_label = ttk.Label(param_frame, text="2.0")
        self.edge_label.grid(row=2, column=2, padx=5)
        edge_slider.configure(command=lambda v: self.edge_label.configure(text=f"{float(v):.1f}"))
        
        # ディテール強調
        ttk.Label(param_frame, text="ディテール:").grid(row=3, column=0, sticky=tk.W, padx=5)
        self.detail_var = tk.DoubleVar(value=2.0)
        detail_slider = ttk.Scale(param_frame, from_=0.0, to=4.0, variable=self.detail_var,
                                 orient=tk.HORIZONTAL, length=200)
        detail_slider.grid(row=3, column=1, padx=5)
        self.detail_label = ttk.Label(param_frame, text="2.0")
        self.detail_label.grid(row=3, column=2, padx=5)
        detail_slider.configure(command=lambda v: self.detail_label.configure(text=f"{float(v):.1f}"))
        
        # コントラスト
        ttk.Label(param_frame, text="コントラスト:").grid(row=4, column=0, sticky=tk.W, padx=5)
        self.contrast_var = tk.DoubleVar(value=1.3)
        contrast_slider = ttk.Scale(param_frame, from_=1.0, to=2.0, variable=self.contrast_var,
                                   orient=tk.HORIZONTAL, length=200)
        contrast_slider.grid(row=4, column=1, padx=5)
        self.contrast_label = ttk.Label(param_frame, text="1.3")
        self.contrast_label.grid(row=4, column=2, padx=5)
        contrast_slider.configure(command=lambda v: self.contrast_label.configure(text=f"{float(v):.1f}"))
        
        # クリアネス（明瞭度）
        ttk.Label(param_frame, text="明瞭度:").grid(row=5, column=0, sticky=tk.W, padx=5)
        self.clarity_var = tk.DoubleVar(value=1.5)
        clarity_slider = ttk.Scale(param_frame, from_=0.0, to=3.0, variable=self.clarity_var,
                                  orient=tk.HORIZONTAL, length=200)
        clarity_slider.grid(row=5, column=1, padx=5)
        self.clarity_label = ttk.Label(param_frame, text="1.5")
        self.clarity_label.grid(row=5, column=2, padx=5)
        clarity_slider.configure(command=lambda v: self.clarity_label.configure(text=f"{float(v):.1f}"))
        
        # 補間方法
        ttk.Label(param_frame, text="補間方法:").grid(row=6, column=0, sticky=tk.W, padx=5)
        self.interp_var = tk.StringVar(value="LANCZOS")
        interp_combo = ttk.Combobox(param_frame, textvariable=self.interp_var, 
                                    values=["LANCZOS", "BICUBIC"],
                                    state="readonly", width=18)
        interp_combo.grid(row=6, column=1, padx=5, sticky=tk.W)
        
        param_frame.columnconfigure(1, weight=1)
        
        # プリセット
        preset_frame = ttk.LabelFrame(main_frame, text="プリセット", padding="10")
        preset_frame.grid(row=2, column=0, columnspan=2, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Button(preset_frame, text="📸 写真モード", 
                  command=lambda: self.apply_preset("photo")).pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_frame, text="📄 文字モード", 
                  command=lambda: self.apply_preset("text")).pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_frame, text="🎨 イラストモード", 
                  command=lambda: self.apply_preset("illustration")).pack(side=tk.LEFT, padx=5)
        ttk.Button(preset_frame, text="💪 最大鮮鋭化", 
                  command=lambda: self.apply_preset("maximum")).pack(side=tk.LEFT, padx=5)
        
        # 画像表示フレーム
        display_frame = ttk.Frame(main_frame)
        display_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # オリジナル画像
        original_label_frame = ttk.LabelFrame(display_frame, text="元画像（ぼやけている）", padding="5")
        original_label_frame.grid(row=0, column=0, padx=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.original_canvas = tk.Canvas(original_label_frame, width=400, height=400, bg='gray')
        self.original_canvas.pack()
        
        # 処理後画像
        enhanced_label_frame = ttk.LabelFrame(display_frame, text="処理後画像（クッキリ！）", padding="5")
        enhanced_label_frame.grid(row=0, column=1, padx=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.enhanced_canvas = tk.Canvas(enhanced_label_frame, width=400, height=400, bg='gray')
        self.enhanced_canvas.pack()
        
        # ステータスバー
        self.status_var = tk.StringVar(value="画像を開いてください")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # グリッド設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        display_frame.columnconfigure(0, weight=1)
        display_frame.columnconfigure(1, weight=1)
        display_frame.rowconfigure(0, weight=1)
        
    def load_image(self):
        """画像を読み込む"""
        file_path = filedialog.askopenfilename(
            title="画像ファイルを選択",
            filetypes=[
                ("画像ファイル", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff"),
                ("すべてのファイル", "*.*")
            ]
        )
        
        if file_path:
            try:
                self.original_image = Image.open(file_path)
                self.enhanced_image = None
                self.display_original()
                self.status_var.set(f"画像を読み込みました: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("エラー", f"画像の読み込みに失敗しました:\n{str(e)}")
                
    def display_original(self):
        """オリジナル画像を表示"""
        if self.original_image:
            img_copy = self.original_image.copy()
            img_copy.thumbnail((400, 400), Image.Resampling.LANCZOS)
            
            photo = ImageTk.PhotoImage(img_copy)
            self.original_canvas.delete("all")
            self.original_canvas.create_image(200, 200, image=photo)
            self.original_canvas.image = photo
            
    def display_enhanced(self):
        """処理後画像を表示"""
        if self.enhanced_image:
            img_copy = self.enhanced_image.copy()
            img_copy.thumbnail((400, 400), Image.Resampling.LANCZOS)
            
            photo = ImageTk.PhotoImage(img_copy)
            self.enhanced_canvas.delete("all")
            self.enhanced_canvas.create_image(200, 200, image=photo)
            self.enhanced_canvas.image = photo
            
    def apply_preset(self, preset_type):
        """プリセットを適用"""
        if preset_type == "photo":
            # 写真モード: バランス重視
            self.scale_var.set(2.0)
            self.sharp_var.set(2.5)
            self.edge_var.set(1.5)
            self.detail_var.set(1.5)
            self.contrast_var.set(1.2)
            self.clarity_var.set(1.3)
            
        elif preset_type == "text":
            # 文字モード: エッジ重視
            self.scale_var.set(2.0)
            self.sharp_var.set(4.0)
            self.edge_var.set(3.0)
            self.detail_var.set(2.5)
            self.contrast_var.set(1.5)
            self.clarity_var.set(2.0)
            
        elif preset_type == "illustration":
            # イラストモード: ディテール重視
            self.scale_var.set(2.5)
            self.sharp_var.set(3.5)
            self.edge_var.set(2.5)
            self.detail_var.set(3.0)
            self.contrast_var.set(1.3)
            self.clarity_var.set(1.8)
            
        elif preset_type == "maximum":
            # 最大鮮鋭化
            self.scale_var.set(2.0)
            self.sharp_var.set(5.0)
            self.edge_var.set(4.0)
            self.detail_var.set(4.0)
            self.contrast_var.set(1.5)
            self.clarity_var.set(3.0)
        
        # ラベルを更新
        self.scale_label.configure(text=f"{self.scale_var.get():.1f}x")
        self.sharp_label.configure(text=f"{self.sharp_var.get():.1f}")
        self.edge_label.configure(text=f"{self.edge_var.get():.1f}")
        self.detail_label.configure(text=f"{self.detail_var.get():.1f}")
        self.contrast_label.configure(text=f"{self.contrast_var.get():.1f}")
        self.clarity_label.configure(text=f"{self.clarity_var.get():.1f}")
        
        self.status_var.set(f"プリセット適用: {preset_type}")
            
    def sharpen_image(self):
        """画像の鮮鋭化処理"""
        if not self.original_image:
            messagebox.showwarning("警告", "画像を読み込んでください")
            return
            
        self.status_var.set("処理中...")
        self.root.update()
        
        try:
            # パラメータ取得
            scale_factor = self.scale_var.get()
            sharpness = self.sharp_var.get()
            edge_enhance = self.edge_var.get()
            detail = self.detail_var.get()
            contrast = self.contrast_var.get()
            clarity = self.clarity_var.get()
            interp_method = self.interp_var.get()
            
            # 補間方法の選択
            if interp_method == "LANCZOS":
                resample = Image.Resampling.LANCZOS
            else:
                resample = Image.Resampling.BICUBIC
            
            img = self.original_image.copy()
            
            # ステップ1: 高品質拡大
            self.status_var.set("画像拡大中...")
            self.root.update()
            new_size = (int(img.width * scale_factor), int(img.height * scale_factor))
            img = img.resize(new_size, resample)
            
            # ステップ2: 強力なアンシャープマスク
            self.status_var.set("鮮鋭化中...")
            self.root.update()
            img = self.apply_strong_unsharp_mask(img, sharpness)
            
            # ステップ3: エッジ強調
            if edge_enhance > 0:
                self.status_var.set("エッジ強調中...")
                self.root.update()
                img = self.enhance_edges(img, edge_enhance)
            
            # ステップ4: ディテール強調
            if detail > 0:
                self.status_var.set("ディテール強調中...")
                self.root.update()
                img = self.enhance_details(img, detail)
            
            # ステップ5: 明瞭度向上（ローカルコントラスト）
            if clarity > 0:
                self.status_var.set("明瞭度向上中...")
                self.root.update()
                img = self.enhance_clarity(img, clarity)
            
            # ステップ6: コントラスト強調
            self.status_var.set("コントラスト調整中...")
            self.root.update()
            enhancer = ImageEnhance.Contrast(img)
            img = enhancer.enhance(contrast)
            
            # ステップ7: 最終シャープニング
            img = img.filter(ImageFilter.SHARPEN)
            
            # ステップ8: アンチエイリアシング（ギザギザ除去）
            # 非常に弱いガウシアンブラーでジャギーだけを除去
            self.status_var.set("ジャギー除去中...")
            self.root.update()
            img = self.remove_jaggies(img)
            
            self.enhanced_image = img
            self.display_enhanced()
            self.status_var.set("処理完了！")
            
        except Exception as e:
            messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{str(e)}")
            self.status_var.set("エラーが発生しました")
            
    def apply_strong_unsharp_mask(self, image, amount):
        """強力なアンシャープマスク"""
        # 複数のブラー半径でアンシャープマスクを適用
        img_array = np.array(image, dtype=np.float32)
        result = img_array.copy()
        
        # 半径1, 2, 3でアンシャープマスク
        for radius in [1, 2, 3]:
            blurred = image.filter(ImageFilter.GaussianBlur(radius=radius))
            blur_array = np.array(blurred, dtype=np.float32)
            
            # アンシャープマスク適用
            mask = img_array - blur_array
            result = result + (amount / 3) * mask
        
        result = np.clip(result, 0, 255)
        return Image.fromarray(result.astype('uint8'))
    
    def enhance_edges(self, image, strength):
        """エッジ強調"""
        # エッジ検出フィルタを複数適用
        img_array = np.array(image, dtype=np.float32)
        
        # SobelフィルタでエッジX
        sobel_x = ndimage.sobel(img_array, axis=0)
        # SobelフィルタでエッジY
        sobel_y = ndimage.sobel(img_array, axis=1)
        
        # エッジの大きさ
        edges = np.sqrt(sobel_x**2 + sobel_y**2)
        
        # エッジを強調して元画像に加算
        enhanced = img_array + strength * 0.2 * edges
        enhanced = np.clip(enhanced, 0, 255)
        
        return Image.fromarray(enhanced.astype('uint8'))
    
    def enhance_details(self, image, strength):
        """ディテール強調（ハイパスフィルタ）"""
        # 大きくぼかした画像を作成
        blurred = image.filter(ImageFilter.GaussianBlur(radius=5))
        
        img_array = np.array(image, dtype=np.float32)
        blur_array = np.array(blurred, dtype=np.float32)
        
        # ハイパスフィルタ（ディテール抽出）
        high_pass = img_array - blur_array
        
        # ディテールを強調して加算
        enhanced = img_array + strength * high_pass
        enhanced = np.clip(enhanced, 0, 255)
        
        return Image.fromarray(enhanced.astype('uint8'))
    
    def enhance_clarity(self, image, strength):
        """明瞭度向上（ローカルコントラスト強調）"""
        # 中程度にぼかした画像を作成
        blurred = image.filter(ImageFilter.GaussianBlur(radius=10))
        
        img_array = np.array(image, dtype=np.float32)
        blur_array = np.array(blurred, dtype=np.float32)
        
        # ローカルコントラストを強調
        local_contrast = img_array - blur_array
        enhanced = img_array + strength * 0.3 * local_contrast
        enhanced = np.clip(enhanced, 0, 255)
        
        return Image.fromarray(enhanced.astype('uint8'))
    
    def remove_jaggies(self, image):
        """ジャギー（ギザギザ）除去"""
        # 非常に弱いブラーでエイリアシングだけを除去
        # エッジは保持しながらギザギザだけを滑らかに
        img_array = np.array(image, dtype=np.float32)
        
        # 極小のガウシアンブラー（σ=0.5）
        if len(img_array.shape) == 3:
            result = np.zeros_like(img_array)
            for i in range(img_array.shape[2]):
                result[:, :, i] = ndimage.gaussian_filter(img_array[:, :, i], sigma=0.5)
        else:
            result = ndimage.gaussian_filter(img_array, sigma=0.5)
        
        return Image.fromarray(result.astype('uint8'))
    
    def save_image(self):
        """処理後の画像を保存"""
        if not self.enhanced_image:
            messagebox.showwarning("警告", "処理後の画像がありません")
            return
            
        file_path = filedialog.asksaveasfilename(
            title="画像を保存",
            defaultextension=".png",
            filetypes=[
                ("PNG画像", "*.png"),
                ("JPEG画像", "*.jpg"),
                ("すべてのファイル", "*.*")
            ]
        )
        
        if file_path:
            try:
                # 最高品質で保存
                if file_path.lower().endswith('.jpg') or file_path.lower().endswith('.jpeg'):
                    self.enhanced_image.save(file_path, quality=98, optimize=True)
                else:
                    self.enhanced_image.save(file_path, optimize=True)
                    
                self.status_var.set(f"保存しました: {os.path.basename(file_path)}")
                messagebox.showinfo("成功", "画像を保存しました")
            except Exception as e:
                messagebox.showerror("エラー", f"保存に失敗しました:\n{str(e)}")
                
    def reset_image(self):
        """画像をリセット"""
        self.original_image = None
        self.enhanced_image = None
        self.original_canvas.delete("all")
        self.enhanced_canvas.delete("all")
        self.status_var.set("画像を開いてください")

def main():
    root = tk.Tk()
    app = ImageSharpenerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
