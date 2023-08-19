import pandas as pd
from scipy.stats import chi2_contingency

pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)

# Excel dosyanızın yolunu belirtin
path = "C:/Users/ŞEYMA/OneDrive/Masaüstü/GotDatas/Game of Thrones All Deaths.xlsx"

# Excel dosyasını okuyarak veriyi bir DataFrame'e yükleyin
df = pd.read_excel(path)

# "None" değerlerini "-" ile değiştirin ve DataFrame'i yeni bir Excel dosyasına kaydedin
df = df.replace(to_replace="None", value=None).fillna("-")
df.to_excel("C:/Users/ŞEYMA/OneDrive/Masaüstü/GotDatas/GotNewData.xlsx", index=False)

# Düzenlenmiş veriyi yükleyin ve analizlerinizi gerçekleştirin
path_new = "C:/Users/ŞEYMA/OneDrive/Masaüstü/GotDatas/GotNewData.xlsx"
df_new = pd.read_excel(path_new)

# Yalnızca ölümün gerçekleştiği bölümleri içeren DataFrame oluşturun
df_olum_gerceklesen_bolumler = df_new[df_new["Death No."].notnull()]

# Toplam ölüm sayısını bulun (yalnızca ölümün gerçekleştiği bölümler üzerinden)
toplam_olum_sayisi = df_olum_gerceklesen_bolumler.shape[0]

# Toplam bölüm sayısını bulun (yalnızca ölümün gerçekleştiği bölümler üzerinden)
toplam_bolum_sayisi = df_olum_gerceklesen_bolumler.groupby("Season")["Episode"].nunique().sum()

# Toplam sezon sayısını bulun (yalnızca ölümün gerçekleştiği bölümler üzerinden)
toplam_sezon_sayisi = df_olum_gerceklesen_bolumler["Season"].nunique()

# Ortalama ölüm sayısını hesaplayın (yalnızca ölümün gerçekleştiği bölümler üzerinden)
ortalama_olum_sayisi = toplam_olum_sayisi / toplam_bolum_sayisi

# Her sezonun toplam ölüm sayısını hesaplayın
olum_sayisi_sezon = df_olum_gerceklesen_bolumler.groupby("Season")["Death No."].count()

# Ortalama ölüm sayısını sezon bazında hesaplayın
ortalama_olum_sayisi_sezon = olum_sayisi_sezon.mean()

print("Toplam Ölüm Sayısı:", toplam_olum_sayisi)
print("Toplam Bölüm Sayısı:", toplam_bolum_sayisi)
print("Toplam Sezon Sayısı:", toplam_sezon_sayisi)
print("Toplam Ölüm Sayıları (Her Sezon İçin):", olum_sayisi_sezon)
print("Ortalama Ölüm Sayısı:", ortalama_olum_sayisi)
print("Ortalama Ölüm Sayısı (Her Sezon İçin):", ortalama_olum_sayisi_sezon)
print(f"En Çok Ölümün Gerçekleştiği Sezon: {olum_sayisi_sezon.idxmax()}, Ölüm Sayısı: {olum_sayisi_sezon.max()}")
print(f"En Az Ölümün Gerçekleştiği Sezon: {olum_sayisi_sezon.idxmin()}, Ölüm Sayısı: {olum_sayisi_sezon.min()}")

# evdeki karakterlerin neden olduğu ölüm sayılarını bulun
ev_olum_sayisi = df_olum_gerceklesen_bolumler["Killers House"].value_counts()
print("Haneler ve neden oldukları ölüm sayıları:", ev_olum_sayisi)
print("En çok ölüme sebep olan hane:", ev_olum_sayisi.idxmax())

print("..............................................................................")

# Hangi hanelerden kaç kişinin öldüğünü bulun 
hane_olum_sayisi = df_olum_gerceklesen_bolumler["Allegiance"].value_counts()
print("haneler ve verdikleri ölü sayısı:", hane_olum_sayisi)
print("en çok ölü veren hane:", hane_olum_sayisi.idxmax())

print("..............................................................................")

# Bağlılık grupları ve lokasyonlar arasındaki ilişkiyi analiz edin
baglilik_lokasyon = df_olum_gerceklesen_bolumler.groupby(["Allegiance", "Location"]).size().reset_index(name='Ölüm Sayısı')
print("Bağlılık Grupları ve Lokasyonlar Arasındaki İlişki:", baglilik_lokasyon)

print("..............................................................................")

# Bağlılık grupları ve yöntemler arasındaki ilişkiyi analiz edin
baglilik_method = df_olum_gerceklesen_bolumler.groupby(["Allegiance", "Method"]).size().reset_index(name='Ölüm Sayısı')
print("Bağlılık Grupları ve Yöntemler Arasındaki İlişki:", baglilik_method)


      ## FREKANS ANALİZİ ##

# Metodların frekans analizini yapın
method_frekans = df_olum_gerceklesen_bolumler["Method"].value_counts()

print("Metodların Frekans Analizi:", method_frekans)
print("En Yaygın Method:", method_frekans.idxmax())

# Lokasyonlara göre frekans analizini yapın
lokasyonlar_frekans = df_olum_gerceklesen_bolumler["Location"].value_counts()

print("Lokasyonlar Frekans Analizi:", lokasyonlar_frekans)
print("En Yaygın Lokasyon:", lokasyonlar_frekans.idxmax())


     ## DAĞILIM ANALİZİ ##
olum_sayisi_sezon_bolum = df_olum_gerceklesen_bolumler.groupby(["Season", "Episode"])["Death No."].count()
print("Ölümlerin Sezon ve Bölümlere Göre Dağılımı:", olum_sayisi_sezon_bolum)


     ## İLİŞKİ ANALİZİ ##
katil_ve_olen = df_olum_gerceklesen_bolumler.groupby(['Killer']).size().reset_index(name='Öldürdüğü Kişi Sayısı')
print("katil ve öldürdüğü kişi sayısı:", katil_ve_olen)

en_cok_olen_katil = katil_ve_olen.loc[katil_ve_olen['Öldürdüğü Kişi Sayısı'].idxmax(), 'Killer']
print("En Çok Kişiyi Öldüren: ", en_cok_olen_katil)


     ## DAĞILIM TESTLERİ ##

# KİLLER_METHOD
# Killer ve Method sütunlarını içeren bir alt veri kümesi oluşturun
katil_ve_yontem = df_new[['Killer', 'Method']]

# Frekans tablosunu oluşturun
frekans_tablosu = pd.crosstab(katil_ve_yontem['Killer'], katil_ve_yontem['Method'])

# Chi-kare testini uygulayın
chi2, p, dof, expected = chi2_contingency(frekans_tablosu)

print("P değeri (Killer ve Method):", p)

# P değeri ile anlamlılığı kontrol edin
if p < 0.05:
    killer_method_iliski = "Killer ve Method arasında anlamlı bir ilişki vardır."
else:
    killer_method_iliski = "Killer ve Method arasında anlamlı bir ilişki yoktur."


# LOCATION_METHOD
# Location ve Method sütunlarını içeren bir alt veri kümesi oluşturun
lokasyon_ve_yontem = df_new[['Location', 'Method']]

# Frekans tablosunu oluşturun
frekans_tablosu = pd.crosstab(lokasyon_ve_yontem['Location'], lokasyon_ve_yontem['Method'])

# Chi-kare testini uygulayın
chi2, p, dof, expected = chi2_contingency(frekans_tablosu)

# P değerini gösterin
print("P değeri (Location ve Method):", p)

# P değeri ile anlamlılığı kontrol edin
if p < 0.05:
    location_method_iliski = "Location ve Method arasında anlamlı bir ilişki vardır."
else:
    location_method_iliski = "Location ve Method arasında anlamlı bir ilişki yoktur."


# Sonuçları içeren bir dictionary oluşturun
sonuclar_dict = {
    "Toplam Ölüm Sayısı": [toplam_olum_sayisi],
    "Toplam Bölüm Sayısı": [toplam_bolum_sayisi],
    "Toplam Sezon Sayısı": [toplam_sezon_sayisi],
    "Toplam Ölüm Sayıları (Her Sezon İçin):": [olum_sayisi_sezon],
    "Ortalama Ölüm Sayısı": [ortalama_olum_sayisi],
    "Ortalama Ölüm Sayısı (Her Sezon İçin)": [ortalama_olum_sayisi_sezon],
    "En Çok Ölümün Gerçekleştiği Sezon": [f"{olum_sayisi_sezon.idxmax()}, Ölüm Sayısı: {olum_sayisi_sezon.max()}"],
    "En Az Ölümün Gerçekleştiği Sezon": [f"{olum_sayisi_sezon.idxmin()}, Ölüm Sayısı: {olum_sayisi_sezon.min()}"],
    "Haneler ve Neden Oldukları Ölüm Sayıları": [ev_olum_sayisi.to_dict()],
    "En Çok Ölüme Sebep Olan Hane": [ev_olum_sayisi.idxmax()],
    "haneler ve verdikleri ölü sayısı:": [hane_olum_sayisi],
    "en çok ölü veren hane:": [hane_olum_sayisi.idxmax()],
    "Bağlılık Grupları ve Lokasyonlar Arasındaki İlişki": [baglilik_lokasyon.to_dict()],
    "Bağlılık Grupları ve Yöntemler Arasındaki İlişki": [baglilik_method.to_dict()],
    "Metodların Frekans Analizi:": [method_frekans],
    "En Yaygın Method:": [method_frekans.idxmax()],
    "Lokasyonlar Frekans Analizi": [lokasyonlar_frekans.to_dict()],
    "En Yaygın Lokasyon": [lokasyonlar_frekans.idxmax()],
    "Ölümlerin Sezon ve Bölümlere Göre Dağılımı": [olum_sayisi_sezon_bolum.to_dict()],
    "Katil ve Öldürdüğü Kişi Sayısı": [katil_ve_olen.to_dict()],
    "En Çok Kişiyi Öldüren": [en_cok_olen_katil],
    "Killer ve Method İlişkisi": [killer_method_iliski],
    "Location ve Method İlişkisi": [location_method_iliski]
}

# Sonuçları içeren DataFrame oluşturun
analiz_sonucları_df = pd.DataFrame(sonuclar_dict)

# Verileri excel dosyasına aktarın
# Ardından analiz sonuçlarını görselleştirmek için ilgili dosyayı Power BI'a aktarabilirsiniz
analiz_sonucları_df.to_excel("AnalizSonucları.xlsx", index=False)
