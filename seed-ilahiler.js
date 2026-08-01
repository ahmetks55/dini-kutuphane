const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require("docx");

const STORAGE = path.join(__dirname, "storage");

function p(text, opts = {}) {
  const runs = [];
  if (opts.rtl) {
    runs.push(
      new TextRun({
        text,
        rtl: true,
        bold: opts.bold,
        color: opts.color || "1a1a1a",
        font: { name: "Traditional Arabic", hint: "eastAsia" },
        size: opts.size || 30,
      })
    );
  } else {
    runs.push(
      new TextRun({
        text,
        bold: opts.bold,
        italics: opts.italics,
        color: opts.color || "1a1a1a",
        size: opts.size || 24,
      })
    );
  }
  return new Paragraph({
    children: runs,
    alignment: opts.rtl ? (opts.align || AlignmentType.RIGHT) : (opts.align || AlignmentType.JUSTIFIED),
    spacing: { after: opts.after != null ? opts.after : 160 },
  });
}

function title(text) {
  return p(text, { bold: true, size: 32, color: "0f5132", align: AlignmentType.CENTER });
}
function section(text) {
  return p(text, { bold: true, size: 26, color: "0f5132" });
}
function blank() {
  return p("", { after: 80 });
}
function flower() {
  return p("❀ ❀ ❀", { align: AlignmentType.CENTER, color: "0f5132", after: 140 });
}

async function writeDocx(filePath, body) {
  const doc = new Document({
    styles: { default: { document: { run: { font: { name: "Segoe UI" } } } } },
    sections: [{ properties: {}, children: body }],
  });
  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(filePath, buf);
}

function makePath(cat, name) {
  return path.join(STORAGE, cat, name + ".docx");
}

async function save(cat, name, body) {
  const fp = makePath(cat, name);
  fs.mkdirSync(path.dirname(fp), { recursive: true });
  if (fs.existsSync(fp)) {
    console.log("  (atlandi, var) " + cat + "/" + name + ".docx");
    return;
  }
  await writeDocx(fp, body);
  console.log("  olusturuldu: " + cat + "/" + name + ".docx");
}

const SERBEST_NOT = "Not: Geleneksel ilâhî metinleri yörelere göre farklılık gösterebilir; cemaatinizde okunan asıl metinle karşılaştırıp kullanmanız tavsiye edilir.";
const NOTALI_NOT = "Notası (porte) resim olarak üretilemediği için bu dosyaya gömülü not eklenmemiştir; ilgili ilâhinin notası TRT nota arşivi ve ilâhî/mevlid nota kaynaklarından temin edilip bu belgeye eklenebilir. Geleneksel ilâhî metinleri yörelere göre değişebilir.";

function ilahiDoc(o) {
  const body = [
    title(o.ad.toUpperCase()),
    p(o.alt || "İlâhî", { italics: true, align: AlignmentType.CENTER, color: "44564e" }),
    blank(),
  ];
  if (o.makam) body.push(p("• Makam: " + o.makam));
  if (o.usul) body.push(p("• Usûl: " + o.usul));
  if (o.besteci) body.push(p("• Bestekâr: " + o.besteci));
  if (o.guftekar) body.push(p("• Güftekâr: " + o.guftekar));
  body.push(blank());
  body.push(flower());
  body.push(section("SÖZLER"));
  body.push(blank());
  (o.sozler || []).forEach((kita, i) => {
    kita.forEach((satir) => body.push(p(satir, { align: AlignmentType.CENTER, after: 40 })));
    body.push(i < o.sozler.length - 1 ? flower() : blank());
  });
  body.push(p(o.not || SERBEST_NOT, { italics: true, color: "44564e" }));
  return body;
}

/* ================= ILÂHİLER ================= */

const ilahiler = [
  /* ----- GENEL (Ilahiler) ----- */
  {
    cat: "Ilahiler", name: "Uyan-Ey-Gozlerim",
    ad: "Uyan Ey Gözlerim", alt: "Gafletten Uyanma İlâhisi", makam: "Eviç",
    besteci: "Anonim (geleneksel)", guftekar: "Yûnus Emre'ye atfedilir",
    sozler: [
      ["Uyan ey gözlerim gafletten uyan", "Uyan uykusu çok gözlerim uyan", "Sabah oldu kuşlar ötüyor", "Uyan uykusu çok gözlerim uyan"],
      ["Günde beş vakit ezan okunur", "Müslüman olan camiye dolunur", "Namaz kılan Hakk'tan bulunur", "Uyan uykusu çok gözlerim uyan"],
      ["Nefis atına binip gaflete daldın", "Gece gündüz günahlara daldın", "Bu dünya fânidir, âhiret gerçek", "Uyan uykusu çok gözlerim uyan"],
    ],
  },
  {
    cat: "Ilahiler", name: "Safa-Dagi",
    ad: "Safâ Dağı", alt: "Bir Çoban Gezer İdi", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Safâ dağı sırtında", "Bir çoban gezer idi", "Resûlün aşkı ile", "Yüreği yanar idi"],
      ["Bir gece melek indi", "Safâ dağına geldi", "Çobana müjde verdi", "Peygamber aşkıyla"],
      ["Çoban sevindi güldü", "Yüzünü sürdü yere", "Şükretti Mevlâ'sına", "Düştü aşk derdine"],
      ["Safâ dağı nurlanır", "Resûl aşkı ile", "Bize de nasip ola", "Şefaat dileğiyle"],
    ],
  },
  {
    cat: "Ilahiler", name: "Allah-Allah-Ilahi",
    ad: "Allah Allah İlâhi", alt: "Yâ Hû — Tevhid İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Allah Allah ilâhi", "Senin adın yâ Hû", "Senin dergâhına kulluk", "Meylimiz yâ Hû"],
      ["Yâ Hû diyelim dâim", "Yâ Hû ile devr edelim", "Hakk'ı zikredelim dâim", "Meylimiz yâ Hû"],
      ["İlâhi sen bağışla", "Günahkâr kullarını", "Rahmetinle kuşat", "Meylimiz yâ Hû"],
    ],
  },
  {
    cat: "Ilahiler", name: "Bir-Gece",
    ad: "Bir Gece", alt: "Melekler Bir Gece", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Bir gece bir gece", "Melekler bir gece", "Semâdan inerler", "Yeryüzüne bir gece"],
      ["Nur ile inerler", "Arş'tan yeryüzüne", "Ümmeti kuşatırlar", "O mübarek gece"],
      ["Sabah olunca giderler", "Kaldırırlar nurunu", "Bize şefaat olsun", "O mübarek gece"],
    ],
  },
  {
    cat: "Ilahiler", name: "Ya-Muhammed-Can-Arzular-Seni",
    ad: "Ya Muhammed Canım Arzular Seni", alt: "Salâvat İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Ya Muhammed canım arzular seni", "İki cihânın mehri olan Resûl", "Cemâlin görmeyen gönül ne bilsin", "Gül cemâlin âşıkları mest eder"],
      ["Her kim seni sevdi ise", "Hak katında makbul oldu", "Ümmetini unutma", "Şefaat eyle bize ya Resûl"],
    ],
  },
  {
    cat: "Ilahiler", name: "Askin-Ile-Asiklar-Yansin",
    ad: "Aşkın İle Âşıklar Yansın", alt: "Yûnus Emre İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Yûnus Emre",
    sozler: [
      ["Aşkın ile âşıklar yansın", "Ben yanmazsam sen yanma", "Ayıplama kardeş beni", "Sen de benim gibi yan"],
      ["Aşkın oduna yanayım", "Derdin deryasına dalayım", "Mecnûn olup ağlayayım", "Sensiz bir dem sürmeyeyim"],
      ["Yûnus senin aşkın ile", "Dertli gönül olup yanar", "Senin aşkın candan tatlı", "Aşkın ile yanmak gerek"],
    ],
  },
  {
    cat: "Ilahiler", name: "Derman-Arardim-Derdime",
    ad: "Derman Arardım Derdime", alt: "Yûnus Emre İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Yûnus Emre",
    sozler: [
      ["Derman arardım derdime", "Derdim bana derman imiş", "Bürhan arardım aslıma", "Aslım bana bürhan imiş"],
      ["Hâr içinde biter gül", "Bülbül konar güle dert", "Yâr ile olmak gerek", "Yârdan ayrılmak neden"],
      ["Ağlarım hatırlarım", "Sırrımı açanlar benim", "Yûnus'un gönlü nurlanır", "Aşkınla yanar gönlüm"],
    ],
  },
  {
    cat: "Ilahiler", name: "Ciktim-Erik-Dalina",
    ad: "Çıktım Erik Dalına", alt: "Yûnus Emre İlâhisi", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Yûnus Emre",
    sozler: [
      ["Çıktım erik dalına", "Orada yedim üzümü", "Bostan issi kakıyıp", "Der ne yersin kozumu"],
      ["Kerpiç koydum kazana", "Poyraz ile kaynattım", "Nedir bu değil diye", "Baltayı kendi sattım"],
      ["Yûnus bir söz söylemiş", "Hiçbir söze benzemez", "Münâfıklar elinden", "Özünü gizler imiş"],
    ],
  },
  {
    cat: "Ilahiler", name: "Bulbulu-Altin-Kafese-Koymuslar",
    ad: "Bülbülü Altın Kafese Koymuşlar", alt: "Yûnus Emre İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Yûnus Emre",
    sozler: [
      ["Bülbülü altın kafese koymuşlar", "İlâhî vatanım demiş", "Beni bir suyun kenarına", "Görüp o yârdan usanmış"],
      ["Döşü altın döşeği", "Ayağında gümüş bağı", "Teninin rengi sararmış", "İlâhî vatanım demiş"],
      ["Yûnus bu sözü söyler", "Miskînâne döker", "Aşk ile yanar gider", "Yâr yoluna düşmüş"],
    ],
  },
  {
    cat: "Ilahiler", name: "Hak-Bir-Gonul-Verdi-Bana",
    ad: "Hak Bir Gönül Verdi Bana", alt: "Yûnus Emre İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Yûnus Emre",
    sozler: [
      ["Hak bir gönül verdi bana", "Hâ dimeden hayran olur", "Bir dem gelir şâd-ı olur", "Bir dem gelir giryân olur"],
      ["Bir dem sanasın kış gibi", "Şol yüzümün rengi soluk", "Bir dem sanasın yaz gibi", "Gülistan ola bu gönül"],
      ["Bir dem söyler sözü tatlı", "Balı yağdan artuk sanur", "Yûnus'un gönlü bu aşkla", "Hak'a gelmiş hayran olur"],
    ],
  },
  {
    cat: "Ilahiler", name: "Sevelim-Sevilelim",
    ad: "Sevelim Sevilelim", alt: "Yûnus Emre İlâhisi", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Yûnus Emre",
    sozler: [
      ["Sevelim sevilelim", "Dünya kimseye kalmaz", "Ölümlü dünya bu", "Bize ne yâd eyler"],
      ["Bu dünyaya gelenler", "Geri gider be kardeş", "Dünya malına aldanma", "Hak için sev be kardeş"],
      ["Yûnus'un sözü budur", "Hakk'ı sevenler güler", "Sevelim sevilelim", "Dünya kimseye kalmaz"],
    ],
  },
  {
    cat: "Ilahiler", name: "Hak-Calabim-Bir-Dil-Verdin",
    ad: "Hak Çalab'ım Bir Dil Verdin", alt: "Yûnus Emre İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Yûnus Emre",
    sozler: [
      ["Hak Çalab'ım bir dil verdin", "Havalara söyler benim", "Ne acebdir bu sırdır", "Can evimde söyler benim"],
      ["Kimse bilmez ne söylerim", "İçerim dışarım bir", "Ne kâfir ne müslim", "Bu sırrı ne söyler benim"],
      ["Yûnus sen bu sırrı açma", "Hak'dan sana derd olur", "Hak Çalab'ım bir dil verdin", "Havalara söyler benim"],
    ],
  },
  {
    cat: "Ilahiler", name: "Bir-Gul-Icin-Bulbul",
    ad: "Bir Gül İçin Bülbül", alt: "Muhabbet İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Bir gül için bülbül gibi", "Gece gündüz ağlarım", "Muhammed'in aşkı ile", "Gönül ateşte yanarım"],
      ["Gül bahçesinde dolaşır", "Bülbül güle ulaşır", "Aşkınla gönül kuşu", "Yâr yüzünü arar"],
      ["Gel ey Resûl'ün ümmeti", "Sev Resûl'ü candan", "Bir gül için bülbül gibi", "Aşkla zikreyle her an"],
    ],
  },
  {
    cat: "Ilahiler", name: "Dun-Gece-Yar-Hanesinde",
    ad: "Dün Gece Yâr Hanesinde", alt: "Aşk İlâhisi", makam: "Acem",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Dün gece yâr hanesinde", "Bir nice seyrân gördüm", "Yâr cemâlin seyrederken", "Kendi hâlimi unuttum"],
      ["Yârın aşkı şulesi", "Gönül evimi sardı", "Dün gece yâr hanesinde", "Gülşen açıldı gönlüme"],
      ["Âşık maşûkun kuludur", "Maşûk âşıka sultan", "Bu gönül yâr yoluna", "Kurban olsun her an"],
    ],
  },
  {
    cat: "Ilahiler", name: "Askin-Derdine-Duseli",
    ad: "Aşkın Derdine Düşeli", alt: "Niyâzî-i Mısrî İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Niyâzî-i Mısrî",
    sozler: [
      ["Aşkın derdine düşeli", "Dünyâ benim olmuş bana", "Bülbül-i gülzâr olalı", "Gülzâr benim olmuş bana"],
      ["Derviş oluban gezeli", "Halkdan yüzüm yumdu çün", "Âşık oluban gezeli", "Dildâr benim olmuş bana"],
      ["Niyâzî bu aşkın ile", "Yandı kül oldu gitti", "Aşkın derdine düşeli", "Dünyâ benim olmuş bana"],
    ],
  },
  {
    cat: "Ilahiler", name: "Aman-Tanrim",
    ad: "Aman Tanrı'm", alt: "Eşrefoğlu Rûmî İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Eşrefoğlu Rûmî",
    sozler: [
      ["Aman Tanrım sen bağışla", "Bizden olan kusurları", "Aman senin dergâhında", "Eriyen erlerden oldum"],
      ["Sana yalvarırım yâ Rab", "Günahım çok, affına muhtacım", "Aman senin dergâhında", "Eriyen erlerden oldum"],
      ["Eşrefoğlu Rûmî düşme sıkılma", "Eyyûb olup sabreyle", "Aman senin dergâhında", "Eriyen erlerden oldum"],
    ],
  },
  {
    cat: "Ilahiler", name: "Ey-Bad-i-Saba",
    ad: "Ey Bâd-ı Sabâ", alt: "Kandil / Medine Özlemi İlâhisi", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Ey bâd-ı sabâ, söyle gül yüzlüme", "Hasretle geçti günlerim", "Resûl'ün kokusu Medine'den", "Gel bâd-ı sabâ getir"],
      ["Ey bâd-ı sabâ varınca", "Yâr eline ulaşınca", "Selâmımı söyle ona", "Resûl'ümün türbesine"],
      ["Sabâ rüzgârı eser", "Gülleri Medine kokar", "Şefaat eylesin bize", "Peygamberim yâ Resûlallah"],
    ],
  },
  {
    cat: "Ilahiler", name: "Mevlid-Gecesi",
    ad: "Mevlid Gecesi", alt: "Kutlu Gece İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Bu gece mübarek gece", "Mevlid gecesi", "Nur iner göklerden", "Âlemlere rahmet gecesi"],
      ["Mustafa doğdu bu gece", "Âlemler nurlandı", "Kâbe'den yükseldi", "Rahmet sesleri"],
      ["Mevlid okunur bugün", "Salâtlar yükselir", "Ümmeti sevinçlidir", "Şefaat gecesi"],
    ],
  },

  /* ----- NOTALI (makam/usûl/besteci bilgisi + nota notu) ----- */
  {
    cat: "Ilahiler/Notali", name: "Sefaat-Ya-Resulallah",
    ad: "Şefaat Ya Resûlallah", alt: "Salâvat İlâhisi", makam: "Rast", usul: "Sofyan",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)", not: NOTALI_NOT,
    sozler: [
      ["Şefaat ya Resûlallah", "Günahım çoktur yâ Resûl", "Kapına geldim yalvarırım", "Şefaat eyle bize"],
      ["Ben kulunum sana geldim", "Elimi açtım duaya", "Ümmetinim ya Resûlallah", "Şefaat eyle bize"],
      ["Yâ Resûlallah seni sevdik", "Aşkın ile yandık", "Mahşerde elimizden tut", "Şefaat eyle bize"],
    ],
  },
  {
    cat: "Ilahiler/Notali", name: "Sultanimizdir-Muhammed",
    ad: "Sultânımızdır Muhammed", alt: "Tevhid İlâhisi", makam: "Rast", usul: "Düyek",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)", not: NOTALI_NOT,
    sozler: [
      ["Sultânımızdır Muhammed", "Rabbimiz Allah bizim", "Hâsulumuzdur Muhammed", "Rabbimiz Allah bizim"],
      ["Yardımcımızdır Muhammed", "Rabbimiz Allah bizim", "Sığınağımızdır Muhammed", "Rabbimiz Allah bizim"],
      ["Âlemlere rahmettir", "Muhammed Mustafa", "Bize şefaatçidir", "Rabbimiz Allah bizim"],
    ],
  },
  {
    cat: "Ilahiler/Notali", name: "Sala-Salati-Ummiye",
    ad: "Salâ — Salât-ı Ümmiyye", alt: "Cuma ve Kandil Salâsı", makam: "Segâh", usul: "Serbest (nöbet)",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)", not: NOTALI_NOT,
    sozler: [
      ["Allâhümme salli alâ seyyidinâ Muhammedin", "Ve alâ âli seyyidinâ Muhammedin"],
      ["Salât-ı dâimeten bi-devâmi mülkillâhi", "El-ekberi'l-ekrem"],
      ["Ve'l-ekremîne'l-mükerremîne", "İlâ yevmi'd-dîni'l-ekber"],
    ],
  },
  {
    cat: "Ilahiler/Notali", name: "La-Ilah-e-Illallah",
    ad: "Lâ İlâhe İllallâh", alt: "Tevhid İlâhisi", makam: "Rast", usul: "Sofyan",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)", not: NOTALI_NOT,
    sozler: [
      ["Lâ ilâhe illallâh", "Muhammedün Resûlullâh", "Zikredelim Hakk'ı dâim", "Lâ ilâhe illallâh"],
      ["Cümle âlem zikreder", "Arş ü kürsî zikreder", "Bülbüller güller ile", "Zikreder Hakk'ı dâim"],
      ["Mü'min olan bu yola", "Baş koyup gelir bugün", "Lâ ilâhe illallâh", "Muhammedün Resûlullâh"],
    ],
  },
  {
    cat: "Ilahiler/Notali", name: "Elhamdulillah-Ya-Resulallah",
    ad: "Elhamdülillâh Ya Resûlallah", alt: "Kandil / Bayram İlâhisi", makam: "Hicaz", usul: "Düyek",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)", not: NOTALI_NOT,
    sozler: [
      ["Elhamdülillâh ya Resûlallah", "Bize seni gönderdi Mevlâ", "Sana ümmet olduk şükrünle", "Elhamdülillâh ya Resûlallah"],
      ["Senin aşkın gönlümüzde", "Salâtlar dilimizde", "Bize şefaat eyle", "Elhamdülillâh ya Resûlallah"],
      ["Kandillerde, bayramlarda", "Mevlidlerde, dualarda", "Anarız seni her an", "Elhamdülillâh ya Resûlallah"],
    ],
  },
  {
    cat: "Ilahiler/Notali", name: "Gul-Yuzlu-Cemaline",
    ad: "Gül Yüzlü Cemâline", alt: "Salâvat İlâhisi", makam: "Eviç", usul: "Düyek",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)", not: NOTALI_NOT,
    sozler: [
      ["Gül yüzlü cemâline", "Hayran oldu gönül", "Cemâlin gülistanında", "Açıldı gül gönlüm"],
      ["Aşkın ile yanarım", "Gül yüzünü ararım", "Şefaatçimsin sen", "Gül yüzlü Resûlüm"],
      ["Cemâline hasretim", "Göster cemâlini", "Gül yüzlü cemâline", "Hayran oldu gönül"],
    ],
  },
  {
    cat: "Ilahiler/Notali", name: "Ya-Resulallah-Seni-Ovmek-Ne-Guzel",
    ad: "Ya Resûlallah Seni Övmek Ne Güzel", alt: "Naat İlâhisi", makam: "Rast", usul: "Sofyan",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)", not: NOTALI_NOT,
    sozler: [
      ["Ya Resûlallah seni övmek ne güzel", "Ne mutlu o âşıklara", "Seni seven ümmetini", "Cennete götürür Resûl"],
      ["Medine'de türben nurlu", "Kokun gülleri andırır", "Ümmetin hasretinle", "Ağlar gece gündüz"],
      ["Ya Resûlallah sana geldim", "Elim boş, yüzüm kara", "Şefaatini dilerim", "Seni övmek ne güzel"],
    ],
  },
  {
    cat: "Ilahiler/Notali", name: "Sumbul-ile-Reyhan",
    ad: "Sümbül İle Reyhan", alt: "Mevlid İlâhisi", makam: "Hüseynî", usul: "Düyek",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)", not: NOTALI_NOT,
    sozler: [
      ["Sümbül ile reyhan ile", "Gülzâr-ı cennet gülü", "Resûl'ün hürmetine", "Açtı bahçeler gönlümde"],
      ["Sümbüller açıldı", "Bülbüller ötüştü", "Resûl'ün aşkı ile", "Güller kokusunu getirdi"],
      ["Açılan her gülde", "Resûl'ü anarız", "Sümbül ile reyhan ile", "Salât ederiz"],
    ],
  },

  /* ----- PEYGAMBERİMİZ ----- */
  {
    cat: "Ilahiler/Peygamberimiz", name: "Cemalin-Gullerini",
    ad: "Cemâlin Güllerini", alt: "Na't-ı Şerîf İlâhisi", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Cemâlin güllerini", "Görüp âşık oldum sana", "Gül yüzünü görmek için", "Canım feda olsun"],
      ["Gül yüzlü Resûlüm", "Gözümün nûrusun", "Cemâlin güllerini", "Bize de göster"],
      ["Aşkın ile yanarım", "Yoluna kurban olurum", "Cemâlin güllerini", "Görmek nasip ola"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Medineye-Selam",
    ad: "Medine'ye Selâm", alt: "Medine Özlemi İlâhisi", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Medine'ye selâm olsun", "Gül kokan yurduna", "Resûl'ün türbesine", "Can-ı gönülden selâm"],
      ["Ey Medine nurlu şehir", "Sana selâm eyledim", "Resûl'ün komşusu olan", "Topraklara selâm"],
      ["Ravza-i Mutahhara'ya", "Akar gözyaşlarım", "Medine'ye varınca", "Selâm söyler gönlüm"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Mubarek-Gecelerde",
    ad: "Mübarek Gecelerde", alt: "Kandil İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Mübarek gecelerde", "Salâtlar yükselir", "Göklerden melekler", "Ümmeti kuşatır"],
      ["Kandil gecelerinde", "Nur yağar yeryüzüne", "Resûl'ün aşkı ile", "Gönüller nurlanır"],
      ["Mübarek gecelerde", "Dualar kabul olur", "Resûl'ün şefaatiyle", "Cennetler müjdelenir"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Peygamberim-Sana-Geldim",
    ad: "Peygamberim Sana Geldim", alt: "Ziyaret / Salâvat İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Peygamberim sana geldim", "El açtım yalvarırım", "Ümmetindenim ben de", "Şefaat eyle bana"],
      ["Yolunu gözledim hep", "Adını andım her an", "Peygamberim sana geldim", "Günahlarımı getirdim"],
      ["Cemâline hasretim", "Ziyaretin özlemi", "Peygamberim sana geldim", "Bizi bırakma mahşerde"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Muhammedin-Askina",
    ad: "Muhammed'in Aşkına", alt: "Mevlid İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Muhammed'in aşkına", "Kâinat var edildi", "Âlemlere rahmet diye", "O gönderildi"],
      ["Muhammed'in aşkına", "Gönüller yandı", "Salât ile selâm", "O'na hep eyledik"],
      ["Muhammed'in aşkına", "Ümmetini bağışla", "Mahşerde şefaat et", "Yâ Rab bu duaya"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Ya-Resulallah-Huzuruna-Geldim",
    ad: "Ya Resûlallah Huzuruna Geldim", alt: "Niyâz İlâhisi", makam: "Segâh",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Ya Resûlallah huzuruna geldim", "Niyâzıma mazhar eyle", "Günahlarımı bağışlat", "Şefaatine nâil eyle"],
      ["Aşkınla yanar gönlüm", "Sana giden yolda", "Ya Resûlallah huzuruna geldim", "Kapında kulluk dilerim"],
      ["Medine'ye gidemezsem", "Adını zikrederim", "Ya Resûlallah huzuruna geldim", "Sen şefaatçi ol bana"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Kutlu-Dogum-Ilahisi",
    ad: "Kutlu Doğum İlâhisi", alt: "Mevlid Kandili İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Kutlu doğum haftasında", "Âlemlere rahmet geldi", "Mustafa doğdu bu günde", "Gönüller nurlandı"],
      ["Melekler tekbir getirdi", "Kâbe'ye nur indi", "Kutlu doğum haftasında", "Salâtlarla anılır"],
      ["Peygamberimizin aşkı", "Her gönülde tazelenir", "Kutlu doğum haftasında", "Rahmet rüzgârları eser"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Sana-Can-Veririm",
    ad: "Sana Cân Veririm", alt: "Muhabbet İlâhisi", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Sana cân veririm cânân yoluna", "Gül yüzünü görmek için", "Resûl'üm aşkına yanar", "Sana cân veririm"],
      ["Gönlüm seninle huzur bulur", "Dilim seninle zikreder", "Sana cân veririm", "Ey Habîb-i Kibriyâ"],
      ["Mahşerde komşu oluruz", "Ümmetine dâhiliz", "Sana cân veririm", "Şefaat eyle bize"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Efendimizin-Kokusu",
    ad: "Efendimizin Kokusu", alt: "Gül Kokusu İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Efendimizin kokusu", "Gül kokusudur", "Gül bahçesinden gelen", "Rahmet rüzgârıdır"],
      ["Gül kokusunu duyunca", "Gönlüm nurla dolar", "Efendimizin kokusu", "Bize şefaat olur"],
      ["Medine'den esen bâd", "Gül kokusu getirir", "Efendimizin kokusu", "Cennet kokusudur"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Ravza-i-Mutahhara",
    ad: "Ravza-i Mutahhara", alt: "Medine / Ravza İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Ravza-i Mutahhara", "Cennet bahçesidir", "Peygamberimizin türbesi", "Gönüllerin sığınağıdır"],
      ["Ravzaya gidenler", "Huzur bulurlar", "Selâm ey Peygamberim", "Ümmetini bağışla"],
      ["Ravza-i Mutahhara", "Nur saçan makamdır", "Orada dua edenler", "Şefaate erer"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Hastayim-Tevhid-Tabibi",
    ad: "Hastayım Tevhid Tabibi", alt: "Salâvat / Na't İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Düşmüşüm dermansız derde", "Yalan hile nefsim sende", "Nasıl bakacağım ben de", "Can Ahmed'in nur yüzüne"],
      ["Hastayım tevhid tabibi", "İns ü cinnin tek sahibi", "Ahmed-i Muhtar habibi", "Nasıl bakarım yüzüne"],
      ["Aşkıyla doğan sabaha", "Banmışım ben çok günaha", "Ne yüzle diyem merhaba", "Can Ahmed'in nur yüzüne"],
      ["Hastayım tevhid tabibi", "İns ü cinnin tek sahibi", "Ahmed-i Muhtar habibi", "Nasıl bakarım yüzüne"],
      ["Viran olmuş kalp ocağım", "Dertlidir gönül kucağım", "Nasıl böyle varacağım", "Can Ahmed'in huzuruna"],
      ["Hastayım tevhid tabibi", "İns ü cinnin tek sahibi", "Ahmed-i Muhtar habibi", "Nasıl bakarım yüzüne"],
      ["Can kurban kudret şahına", "O enbiya sultanına", "Nasıl giderim yanına", "Can Ahmed'in nur yüzüne"],
      ["Hastayım tevhid tabibi", "İns ü cinnin tek sahibi", "Ahmed-i Muhtar habibi", "Nasıl bakarım yüzüne"],
      ["Sefil hâlinle özüne", "Girmişsin nefsin közüne", "Nasıl bakarım mahşerde", "Can Ahmed'in nur yüzüne"],
      ["Hastayım tevhid tabibi", "İns ü cinnin tek sahibi", "Ahmed-i Muhtar habibi", "Nasıl bakarım yüzüne"],
    ],
  },
  {
    cat: "Ilahiler/Peygamberimiz", name: "Askin-Ile-Asiklar-Ya-Resulallah",
    ad: "Aşkın İle Âşıklar Yansın Ya Rasûlallah", alt: "Na't-ı Şerîf İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Yûnus Emre",
    sozler: [
      ["Aşkın ile aşıklar", "Yansın ya Rasûlallah", "İçip aşkın şerabın", "Kansın ya Rasûlallah"],
      ["Âşık oldum dildâre", "Bülbülüm şol gülzâre", "Seni sevmeyen nâre", "Yansın ya Rasûlallah"],
      ["Şol seni seven kişi", "Verir yoluna başı", "İki cihan güneşi", "Sensin ya Rasûlallah"],
      ["Şol seni sevenlere", "Kıl şefaat onlara", "Mümin olan tenlere", "Cansın ya Rasûlallah"],
      ["Şol seni seven Sübhan", "Oldu kamuya sultan", "Canım yoluna kurban", "Olsun ya Rasûlallah"],
      ["Âşık Yûnus'un canı", "İlm-ü şefaat kânı", "Âlemlerin sultanı", "Sensin ya Rasûlallah"],
    ],
  },

  /* ----- HAC ----- */
  {
    cat: "Ilahiler/Hac", name: "Leblebbeyk",
    ad: "Leblebbeyk", alt: "Telbiye İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Lebbeyk Allâhümme Lebbeyk", "Lebbeyk lâ şerîke leke lebbeyk", "İnne'l-hamde ve'n-ni'mete", "Leke ve'l-mülk, lâ şerîke lek"],
      ["Beytullâh'a varıyoruz", "Hacı olduk inşallah", "Lebbeyk diye koşuyoruz", "Mevlâ'ya yalvarıyoruz"],
      ["Mîkâtta ihram giydik", "Telbiyeyi getirdik", "Lebbeyk Allâhümme Lebbeyk", "Rabbimiz duâmızı kabul eyle"],
    ],
  },
  {
    cat: "Ilahiler/Hac", name: "Beytullaha",
    ad: "Beytullâh'a", alt: "Kâbe Özlemi İlâhisi", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Beytullâh'a varınca", "Yüzümü sürdüm taşına", "Kâbe'ye giden yolda", "Gönlümü verdim Hakk'a"],
      ["Beytullâh'a kavuştum", "Mevlâ'ma yalvardım", "Zemzem içtim ağladım", "Günahlarıma tövbe ettim"],
      ["Tavaf eder Kâbe'yi", "Mü'minler huzurla", "Beytullâh'a varanlar", "Cennete yaklaşır"],
    ],
  },
  {
    cat: "Ilahiler/Hac", name: "Arafatta",
    ad: "Arafat'ta", alt: "Vakfe İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Arafat'ta vakfe durdum", "Elimi açtım Mevlâ'ya", "Günahlarımdan utandım", "Bağışla yâ Rab dedim"],
      ["Arafat'ta günahkârlar", "Göz yaşlarıyla tövbe eder", "Rahmet deryası coşar", "Hacılar huzur bulur"],
      ["Müzdelife'ye indik", "Şeytanı taşladık", "Arafat'ta vakfe ile", "Hacılık tamam oldu"],
    ],
  },
  {
    cat: "Ilahiler/Hac", name: "Kabe-Yollarinda",
    ad: "Kâbe Yollarında", alt: "Hac Yolu İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Kâbe yollarında", "Yürür hacılar", "Lebbeyk diye diye", "Akarlar Kâbe'ye"],
      ["Kâbe yollarında", "Nurlar saçılır", "Resûl'ün şehrine", "Selâmlar gönderilir"],
      ["Kâbe'ye varınca", "Gönüller yıkanır", "Kâbe yollarında", "Dualar kabul olur"],
    ],
  },
  {
    cat: "Ilahiler/Hac", name: "Haci-Ya-Resulallah",
    ad: "Hacı Ya Resûlallah", alt: "Medine Ziyareti İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Hacı ya Resûlallah", "Ziyarete geldim", "Senin aşkın ile", "Yollara düştüm"],
      ["Ravzana varınca", "Gözlerim dolar", "Hacı ya Resûlallah", "Selâm olsun sana"],
      ["Ümmetindenim ben", "Şefaatinle geldim", "Hacı ya Resûlallah", "Kapında kulunum"],
    ],
  },
  {
    cat: "Ilahiler/Hac", name: "Zemzem",
    ad: "Zemzem", alt: "Zemzem Suyu İlâhisi", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Zemzem suyu Kâbe'nin", "Sofrasıdır mü'minin", "Zemzem içen hastalar", "Şifa bulur derler"],
      ["Zemzem deryası coşar", "Gönüllere ferahlık verir", "Hacılar içip şükreder", "Rabbimize hamd eder"],
      ["Zemzem suyu mübarek", "Rahmetten inmiştir", "Zemzem içen âşıklar", "Kâbe'ye selâm eder"],
    ],
  },

  /* ----- ÖLÜM ----- */
  {
    cat: "Ilahiler/Olum", name: "Aglar-Isem-Simdi",
    ad: "Ağlar İsem Şimdi", alt: "Rihlet (Vefat) İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Ağlar isem şimdi, ağlamaz mıyım", "Resûl'üm gitti bu dünyadan", "Gül yüzünü göremeden", "Kaldım âvâre bu dünyada"],
      ["Rihlet günü geldi çattı", "Âlemler yas tuttu", "Resûl'üm ayrıldı gitti", "Ümmeti ağlar yollarda"],
      ["Ağlar isem şimdi, ağlamaz mıyım", "Sahâbeler yas tuttu", "Resûl'ümün ayrılığı", "Yüreğimize dağ oldu"],
    ],
  },
  {
    cat: "Ilahiler/Olum", name: "Olum-Var",
    ad: "Ölüm Var", alt: "Ölümü Hatırlatma İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Ölüm var, ölüm var", "Her nefeste ölüm var", "Bugün bize ne gelir", "Yarın bize ölüm var"],
      ["Dünya fâni bir süs", "Gönül ona verme", "Ölüm var, ölüm var", "Âhirete göç var"],
      ["Ölümü hatırla", "Gafletten uyan", "Ölüm var, ölüm var", "Her nefeste ölüm var"],
    ],
  },
  {
    cat: "Ilahiler/Olum", name: "Bu-Dunya-Bir-Pencere",
    ad: "Bu Dünya Bir Pencere", alt: "İbret İlâhisi", makam: "Hicaz",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Bu dünya bir pencere", "Bir bakılan yerdir", "Herkes bir gün çıkar", "Bu dünya fânidir"],
      ["Bu dünyada ne kazandın", "Âhirete ne götürdün", "Bu dünya bir pencere", "Bir bakılan yerdir"],
      ["Sakın aldanma dünyaya", "O yalan söyler", "Bu dünya bir pencere", "Bir bakılan yerdir"],
    ],
  },
  {
    cat: "Ilahiler/Olum", name: "Ey-Gonul-Ibret-Al",
    ad: "Ey Gönül İbret Al", alt: "Ölüm Tefekkürü İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Ey gönül ibret al", "Şu ölümden", "Her canlı ölümü tadar", "Sana da gelecek"],
      ["Kabir karanlık oda", "Amel nur olsun", "Ey gönül ibret al", "Şu ölümden"],
      ["Gafletten uyan gönül", "Hazırlığını yap", "Ey gönül ibret al", "Şu ölümden"],
    ],
  },
  {
    cat: "Ilahiler/Olum", name: "Yol-Gorundu",
    ad: "Yol Göründü", alt: "Âhiret Yolu İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Yol göründü, yol göründü", "Âhiret yolu göründü", "Dünya fâni, yol budur", "Âhiret yolu göründü"],
      ["Gözüm yumdum bu dünyadan", "Gitti gençlik, geldi ihtiyarlık", "Yol göründü, yol göründü", "Âhiret yolu göründü"],
      ["Ölüm geldi kapıma", "Ecel yazdı başıma", "Yol göründü, yol göründü", "Âhiret yolu göründü"],
    ],
  },
  {
    cat: "Ilahiler/Olum", name: "Elveda-Elveda-Elveda-Dunyam",
    ad: "Elveda Elveda Elveda Dünyam", alt: "Ölüm / Dünyaya Veda İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Bugün benim ölüm günüm", "Elveda elveda elveda dünyam", "Ne dün kaldı ne bugünüm", "Elveda elveda elveda dünyam"],
      ["Ağlamasın anam bacım", "İman olsun başta tacım", "Ne dert kaldı ne de acım", "Elveda elveda elveda dünyam"],
      ["Ölüm haberim duyuldu", "Suyum kazana koyuldu", "Dünya libasım soyuldu", "Elveda elveda elveda dünyam"],
      ["Ağlamasın anam bacım", "İman olsun başta tacım", "Ne dert kaldı ne de acım", "Elveda elveda elveda dünyam"],
      ["Beyaz libas giydirdiler", "Tahta ata bindirdiler", "Toprak eve indirdiler", "Elveda elveda elveda dünyam"],
      ["Ağlamasın anam bacım", "İman olsun başta tacım", "Ne dert kaldı ne de acım", "Elveda elveda elveda dünyam"],
    ],
  },
  {
    cat: "Ilahiler/Olum", name: "Kevser-Irmaginda-Saki-Olan-Yar",
    ad: "Kevser Irmağında Saki Olan Yar", alt: "Âhiret / Sırat İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Aşık Sefil Selimi (Ahmet Günbulut)",
    sozler: [
      ["Kevser ırmağında saki olan yar", "Bir bardak dem ikram etmez mi ola", "Sıratın yolunu iyi bilen yar", "Benim de elimden tutmaz mı ola"],
      ["Aman medet duy sesimi dardayım", "Sorma hallerimi gayet zordayım", "Cehennemden daha beter kordayım", "Yanarım yandığım yetmez mi ola"],
      ["Her yanımı harlı ateş çevirdi", "Vücut sarayımı yaktı kavurdu", "Yaptım mamur ettim geri devirdi", "Viranemde güller bitmez mi ola"],
      ["Zindanda düşsem de gam yemem yine", "Sefil Selimi'yle dursan yan yana", "Olmak istiyorum dostla can cana", "Muradımca bülbül ötmez mi ola"],
    ],
  },

  /* ----- ÇOCUK ----- */
  {
    cat: "Ilahiler/Cocuk", name: "Guldur-Gulum-Guldur",
    ad: "Güldür Gülüm Güldür", alt: "Çocuk İlâhisi", makam: "Hüseynî",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Güldür gülüm güldür", "Sevimli gülüm güldür", "Allah'ım sen güldür", "Bu güzel günümüzde"],
      ["Gülüm açılsın gülsün", "Kalbimiz sevinç dolsun", "Güldür gülüm güldür", "Bu mübarek günümüzde"],
      ["Bayramımız kutlu olsun", "Sevinçler çoğalsın", "Güldür gülüm güldür", "Bu güzel günümüzde"],
    ],
  },
  {
    cat: "Ilahiler/Cocuk", name: "Ne-Guzel-Yaratmissin",
    ad: "Ne Güzel Yaratmışsın", alt: "Çocuk Şükür İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Ne güzel yaratmışsın", "Sen her şeyi yâ Rabbim", "Güneşi, ayı, yıldızı", "Yarattın bizler için"],
      ["Gökyüzü mavi mavi", "Kuşlar cıvıl cıvıl", "Ne güzel yaratmışsın", "Sen her şeyi yâ Rabbim"],
      ["Annemi, babamı verdin", "Bana bu güzel dünyayı", "Sana şükürler olsun", "Ne güzel yaratmışsın"],
    ],
  },
  {
    cat: "Ilahiler/Cocuk", name: "Allahim-Biliyorum",
    ad: "Allah'ım Biliyorum", alt: "Çocuk Dua İlâhisi", makam: "Rast",
    besteci: "Anonim (geleneksel)", guftekar: "Anonim (geleneksel)",
    sozler: [
      ["Allah'ım biliyorum", "Sen her şeyi görürsün", "Kalbimdeki duayı", "Sen daha iyi bilirsin"],
      ["Sevgiyle dolu kalbim", "Seviyorum herkesi", "Allah'ım biliyorum", "Sen bizi seversin"],
      ["Namazımı kılarım", "Sana şükrederim", "Allah'ım biliyorum", "Sen her şeyi görürsün"],
    ],
  },
];

/* ================= URETIM ================= */

(async () => {
  console.log("Ilahiler uretiliyor...\n");
  let sayac = 0;
  for (const ilahi of ilahiler) {
    const body = ilahiDoc(ilahi);
    await save(ilahi.cat, ilahi.name, body);
    sayac++;
  }
  console.log("\nToplam ilâhi dosyasi: " + sayac);
  console.log("Tamamlandi.");
})();
