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

/* ================= ICERIKLER ================= */

const programDuasi = [
  title("PROGRAM DUASI"),
  section("(Mevlid / Cemiyet Duasi)"),
  blank(),
  p("Bismillâhirrahmânirrahîm."),
  p("Elhamdülillâhi rabbil'âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn."),
  p("Lâ havle velâ kuvvete illâ billâhil-aliyyil-azîm."),
  blank(),
  p("Yâ Rabbe'l-âlemîn! Bu mübarek mecliste toplanan, birbirini seven, sevdiren mü'min kardeşlerimizin kalplerini nûr, yüzlerini ferah, gönüllerini sürûr eyle. Yüzlerimizi dîdâr-ı kerîminle, kalplerimizi muhabbet-i ilâhiyyenle doldur. İmânımızı, kitabımızı, Kur'ân'ımızı, Peygamberimizi bize mahşer gününde şefaatçi eyle."),
  p("Yâ Rab! Sevgili Peygamberimiz Hazret-i Muhammed Mustafâ'nın (sallallâhü aleyhi ve sellem) hürmetine, Kur'ân-ı azîmüşşânın hürmetine, bu mecliste okunan âyet-i kerîmelerin, salevât-ı şerîfelerin ve mevlid-i şerîfin hürmetine; üzerimizde emeği bulunan annelerimizi, babalarımızı, hocalarımızı, büyüklerimizi cümlesini râzı eyle. Ölmüşlerimize rahmet, hastalarımıza şifâ, dertlilerimize devâ, borçlularımıza edâ, gariplere vatan, tutuklulara serbestlik, yolculara selâmet ihsan eyle."),
  p("Yâ Rab! Devletimize, milletimize, vatanımıza, ezânımıza, bayrağımıza, dinimize ve diyânetimize zeval verme. Düşmanlarımıza fırsat verme. Âlem-i İslâm'ın birliğini, dirliğini ve düzenini dâim eyle. Mazlum ve mağdur Müslüman kardeşlerimize zafer ve feth-i mübîn ihsan eyle."),
  p("Yâ Rab! Okuduğumuz aşr-ı şerîfleri, salevât-ı şerîfeleri ve mevlid-i şerîfi dergâh-ı izzetinde kabul buyur. Sevabını başta Peygamberlerin rûhlarına, ashâb-ı güzîne, ehl-i beyt-i kirâma, cümle şühedâ ve gazilere ve bütün mü'min ve müslimlerin ölmüşlerinin rûhlarına hediye eyledik; kabul eyle yâ Rabbe'l-âlemîn."),
  p("Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn."),
];

const hatimDuasi = [
  title("HATİM DUASI"),
  blank(),
  p("Elhamdülillâhi rabbil'âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn."),
  p("Allâhümme'tağfir li'l-mü'minîne ve'l-mü'minât ve'l-müslimîne ve'l-müslimât, el-ahyâi minhüm ve'l-emvât. İnneke alâ külli şey'in kadîr."),
  blank(),
  p("Yâ Rabbe'l-âlemîn! Bizlere Kur'ân-ı azîmüşşânı okuyup hatmetmeyi nasip eyledin; şükürler olsun. Bu hatm-i şerîfin sevâbını, başta sevgili Peygamberimiz Hazret-i Muhammed Mustafâ'nın (s.a.v.) rûh-u şerîflerine, dört büyük halîfemize, ehl-i beyt-i kirâma, ashâb-ı güzîne, evliyâ ve enbiyâya, cümle şühedâ ve gazilere, rûhları şehid ve gazilerin rûhlarına karışmış olan bütün geçmişlerimizin rûhlarına ve hassaten bu meclisimizde rûhları huzûrumuzda hazır bulunan cümle ölmüşlerimizin rûhlarına hediye eyledik; kabul eyle yâ Rabbe'l-âlemîn."),
  p("Yâ Rab! Bu hatmin hürmetine; hasta olanlarımıza şifâ, borçlu olanlarımıza edâ, işsizlerimize iş, evlenmek isteyenlere hayırlı nasib, çocuğu olmayanlara sâlih evlât nasip eyle. Din ve vatan uğrunda şehid düşen kardeşlerimize rahmet, ailelerine sabr-ı cemîl, metanet ve sabırlar ihsan eyle."),
  p("Yâ Rab! Bu duayı okuyanların ve dinleyenlerin, bu hatme iştirak edenlerin hepsinin cümle günahlarını mağfiret eyle. Hastalara şifâ, dertlilere devâ eyle. Bütün mü'minlerin dualarını dergâh-ı izzetinde kabul ve makbul eyle."),
  p("Âmîn, bi-hürmet-i seyyidi'l-mürselîn. Ve'l-hamdü lillâhi rabbi'l-âlemîn."),
];

function duaDoc(titleText, subtitle, entries) {
  const body = [title(titleText), p(subtitle, { italics: true, align: AlignmentType.CENTER, color: "44564e" }), blank()];
  entries.forEach((e) => {
    if (e.section) body.push(section(e.section));
    if (e.ar) body.push(p(e.ar, { rtl: true, after: 120 }));
    if (e.okunus) body.push(p(e.okunus, { italics: true, color: "44564e", after: 120 }));
    if (e.meal) body.push(p(e.meal, { italics: true, color: "44564e" }));
    body.push(blank());
  });
  return body;
}

const vaazBaslama = duaDoc("VAAZ BAŞLAMA DUASI", "(Vaaz, sohbet, ders ve hatim programları öncesi okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm." },
  { ar: "الْحَمْدُ لِلَّهِ رَبِّ الْعَالَمِينَ وَالصَّلَاةُ وَالسَّلَامُ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Elhamdülillâhi rabbi'l-âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Hamd, âlemlerin Rabbi Allah'a mahsustur. Salât ve selâm, efendimiz Muhammed'in, âlinin ve ashâbının üzerine olsun." },
  { ar: "اللَّهُمَّ صَلِّ عَلَى مُحَمَّدٍ وَعَلَى آلِ مُحَمَّدٍ كَمَا صَلَّيْتَ عَلَى إِبْرَاهِيمَ وَعَلَى آلِ إِبْرَاهِيمَ إِنَّكَ حَمِيدٌ مَجِيدٌ", okunus: "Allâhümme salli alâ Muhammedin ve alâ âli Muhammed kemâ salleyte alâ İbrâhîme ve alâ âli İbrâhîm. İnneke hamîdün mecîd.", meal: "Allah'ım! İbrâhîm'e ve İbrâhîm'in âline rahmet ettiğin gibi Muhammed'e ve Muhammed'in âline de rahmet eyle. Şüphesiz sen övülmeye lâyık, şanı yüce olansın." },
  { section: "Rabbim, göğsümü aç, işimi kolaylaştır (Tâhâ 25-28)" },
  { ar: "رَبِّ اشْرَحْ لِي صَدْرِي وَيَسِّرْ لِي أَمْرِي وَاحْلُلْ عُقْدَةً مِنْ لِسَانِي يَفْقَهُوا قَوْلِي", okunus: "Rabbişrah lî sadrî ve yessir lî emrî vehlül ukdeten min lisânî yefkahû kavlî.", meal: "Rabbim! Gönlüme genişlik ver, işimi kolaylaştır, dilimden düğümü çöz ki sözümü anlasınlar." },
  { section: "Rabbim, ilmimi artır (Tâhâ 114)" },
  { ar: "رَبِّ زِدْنِي عِلْمًا", okunus: "Rabbi zidnî ilmâ.", meal: "Rabbim! Benim ilmimi artır." },
  { section: "Rabbim, şükür ve sâlih amel nasip eyle (Neml 19)" },
  { ar: "رَبِّ أَوْزِعْنِي أَنْ أَشْكُرَ نِعْمَتَكَ الَّتِي أَنْعَمْتَ عَلَيَّ وَعَلَى وَالِدَيَّ وَأَنْ أَعْمَلَ صَالِحًا تَرْضَاهُ وَأَدْخِلْنِي بِرَحْمَتِكَ فِي عِبَادِكَ الصَّالِحِينَ", okunus: "Rabbi evzi'nî en eşküra ni'meteke'lletî en'amte aleyye ve alâ vâlideyye ve en a'mele sâlihan terdâhu ve edhılnî bi-rahmetike fî ibâdike's-sâlihîn.", meal: "Rabbim! Bana ve anne-babama verdiğin nimete şükretmeyi, razı olacağın sâlih işler yapmayı bana nasip et; rahmetinle beni sâlih kullarının arasına kat." },
  { section: "Rabbimiz, bize dünya ve âhiret iyiliği ver (Bakara 201)" },
  { ar: "رَبَّنَا آتِنَا فِي الدُّنْيَا حَسَنَةً وَفِي الْآخِرَةِ حَسَنَةً وَقِنَا عَذَابَ النَّارِ", okunus: "Rabbenâ âtinâ fi'd-dünyâ haseneten ve fi'l-âhirati haseneten ve kınâ azâbe'n-nâr.", meal: "Rabbimiz! Bize dünyada da iyilik ver, âhirette de iyilik ver ve bizi cehennem azabından koru." },
  { section: "Kapanış (Türkçe hatim üslubuyla dua)" },
  { meal: "Yâ Rab! Vaazlarımızı, sohbetlerimizi ve okuduğumuz bu âyetleri dergâh-ı izzetinde kabul eyle. İlmimizi artır, hikmetimizi çoğalt. Kalplerimizi Kur'ân ve sünnet nûruyla aydınlat; dillerimizi hak söylemek için, gönüllerimizi hakikati anlamak için aç. Bize ve dinleyenlerimize hidayet nasip eyle. Sözlerimizi ve amellerimizi sırf senin rızan için ihlâslı kıl. Okuduğumuz bu hatmin ve yaptığımız bu duanın sevabını başta sevgili Peygamberimiz (s.a.v.) olmak üzere bütün peygamberlere, ehl-i beyte, ashâb-ı güzîne ve tüm inananların rûhlarına hediye eyledik; kabul eyle. Şu anda burada hazır bulunan ve huzurlarına erişemeyen bütün kardeşlerimizi, âilemizi, evlâtlarımızı, hastalarımızı, borçlularımızı ve yokluk içinde olan kullarını rahmetinle kuşat. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const cenazeErkek = duaDoc("CENAZE DUASI — ERKEK", "(Erkek bir mü'minin cenaze namazında okunur; hatim programı kapsamında vefat edenler için de okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "إِنَّا لِلَّهِ وَإِنَّا إِلَيْهِ رَاجِعُونَ", okunus: "İnnâ lillâhi ve innâ ileyhi râciûn.", meal: "Biz Allah'a aitiz ve şüphesiz O'na döneceğiz. (Bakara 156)" },
  { ar: "اللَّهُمَّ صَلِّ عَلَى مُحَمَّدٍ وَعَلَى آلِ مُحَمَّدٍ كَمَا صَلَّيْتَ عَلَى إِبْرَاهِيمَ وَعَلَى آلِ إِبْرَاهِيمَ إِنَّكَ حَمِيدٌ مَجِيدٌ", okunus: "Allâhümme salli alâ Muhammedin ve alâ âli Muhammed kemâ salleyte alâ İbrâhîme ve alâ âli İbrâhîm. İnneke hamîdün mecîd.", meal: "Allah'ım! İbrâhîm'e ve İbrâhîm'in âline rahmet ettiğin gibi Muhammed'e ve Muhammed'in âline de rahmet eyle. Şüphesiz sen övülmeye lâyık, şanı yüce olansın." },
  { section: "Cenaze Duası" },
  { ar: "اللَّهُمَّ اغْفِرْ لِحَيِّنَا وَمَيِّتِنَا وَشَاهِدِنَا وَغَائِبِنَا وَصَغِيرِنَا وَكَبِيرِنَا وَذَكَرِنَا وَأُنْثَانَا", okunus: "Allâhümme'ğfir li-hayyinâ ve meyyitinâ ve şâhidinâ ve ğâibinâ ve sağîrinâ ve kebîrinâ ve zekerinâ ve ünsânâ.", meal: "Allah'ım! Dirimizi, ölümüzü, hazır bulunanımızı, gâibimizi, küçüğümüzü, büyüğümüzü, erkeğimizi ve kadınımızı bağışla." },
  { ar: "اللَّهُمَّ مَنْ أَحْيَيْتَهُ مِنَّا فَأَحْيِهِ عَلَى الْإِسْلَامِ وَمَنْ تَوَفَّيْتَهُ مِنَّا فَتَوَفَّهُ عَلَى الْإِيمَانِ", okunus: "Allâhümme men ahyeytehû minnâ fe-ahyihi ale'l-İslâmi ve men teveffeytehû minnâ feteveffehû ale'l-îmân.", meal: "Allah'ım! Bizden yaşattığını İslâm üzere yaşat, ecelini verdiğini de iman üzere vefat ettir." },
  { ar: "اللَّهُمَّ اغْفِرْ لَهُ وَارْحَمْهُ وَعَافِهِ وَاعْفُ عَنْهُ وَأَكْرِمْ نُزُلَهُ وَوَسِّعْ مُدْخَلَهُ وَاغْسِلْهُ بِالْمَاءِ وَالثَّلْجِ وَالْبَرَدِ وَنَقِّهِ مِنَ الْخَطَايَا كَمَا يُنَقَّى الثَّوْبُ الْأَبْيَضُ مِنَ الدَّنَسِ", okunus: "Allâhümme'ğfir lehû verhamhû ve âfihî ve'fü anhû ve ekrim nüzülehû ve vessi' medhalehû ve'ğsilhû bi'l-mâi ve's-selci ve'l-bered ve nekkihî mine'l-hatâyâ kemâ yünakkâ's-sevbü'l-ebyadu mine'd-denes.", meal: "Allah'ım! Onu bağışla, ona merhamet et, ona afiyet ver, onu affet, konduğu yeri şereflendir, kabrini genişlet; onu su, kar ve dolu ile yıka, beyaz elbisenin kirden temizlendiği gibi onu hatalardan temizle." },
  { ar: "وَأَبْدِلْهُ دَارًا خَيْرًا مِنْ دَارِهِ وَأَهْلًا خَيْرًا مِنْ أَهْلِهِ وَزَوْجًا خَيْرًا مِنْ زَوْجِهِ وَأَدْخِلْهُ الْجَنَّةَ وَأَعِذْهُ مِنْ عَذَابِ الْقَبْرِ وَعَذَابِ النَّارِ", okunus: "Ve ebdilhû dâran hayran min dârihî ve ehlen hayran min ehlihî ve zevcen hayran min zevcihî ve edhılhü'l-cennete ve a'izhû min azâbi'l-kabri ve azâbi'n-nâr.", meal: "Ona, bulunduğu evden daha hayırlı bir ev, daha hayırlı aile ve eş ver; onu cennete koy ve kabir ile cehennem azabından koru." },
  { section: "Kabir Sorularında Sabit Kılma Duası" },
  { ar: "اللَّهُمَّ ثَبِّتْهُ بِالْقَوْلِ الثَّابِتِ", okunus: "Allâhümme sebbithû bi'l-kavli's-sâbit.", meal: "Allah'ım! Onu, dünyada da âhirette de sabit söz (kelime-i şehâdet) üzere sabit kıl." },
  { ar: "اللَّهُمَّ نَوِّرْ قَبْرَهُ وَوَسِّعْ مَدْخَلَهُ وَآمِنْهُ مِنْ فَزَعِ الْيَوْمِ الْأَكْبَرِ", okunus: "Allâhümmе nûvvir kabrahu ve vessi' medhalehû ve âminhü min fezei'l-yevmi'l-ekber.", meal: "Allah'ım! Kabrini nurlandır, girişini genişlet ve onu büyük kıyamet gününün dehşetinden emin kıl." },
  { section: "Hatim Programı Kapsamında Kapanış" },
  { meal: "Yâ Rab! Okuduğumuz bu hatmin, yaptığımız bu duanın ve cenaze namazını kılanların dualarının sevabını, başta bu vefat eden kardeşimiz olmak üzere bütün mü'min erkek ve kadınların rûhlarına hediye eyledik; kabul eyle. Yâ Rab! Ahiret yolculuğuna çıkan bu kulunu sen bağışla, ona rahmet eyle. Onu sabit söz üzere sabit kıl, kabrini cennet bahçelerinden bir bahçe eyle. Ardında kalan ailesine, yakınlarına ve cenazesinde dua edenlere sabır, tesellî ve ecr-i azîm nasip eyle. Bizleri de onu hayırla yâd edenlerden, ölümü tefekkür edip kendini hazırlayanlardan eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const cenazeKadin = duaDoc("CENAZE DUASI — KADIN", "(Kadın bir mü'mine'nin cenaze namazında okunur; hatim programı kapsamında vefat edenler için de okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "إِنَّا لِلَّهِ وَإِنَّا إِلَيْهِ رَاجِعُونَ", okunus: "İnnâ lillâhi ve innâ ileyhi râciûn.", meal: "Biz Allah'a aitiz ve şüphesiz O'na döneceğiz. (Bakara 156)" },
  { ar: "اللَّهُمَّ صَلِّ عَلَى مُحَمَّدٍ وَعَلَى آلِ مُحَمَّدٍ كَمَا صَلَّيْتَ عَلَى إِبْرَاهِيمَ وَعَلَى آلِ إِبْرَاهِيمَ إِنَّكَ حَمِيدٌ مَجِيدٌ", okunus: "Allâhümme salli alâ Muhammedin ve alâ âli Muhammed kemâ salleyte alâ İbrâhîme ve alâ âli İbrâhîm. İnneke hamîdün mecîd.", meal: "Allah'ım! İbrâhîm'e ve İbrâhîm'in âline rahmet ettiğin gibi Muhammed'e ve Muhammed'in âline de rahmet eyle. Şüphesiz sen övülmeye lâyık, şanı yüce olansın." },
  { section: "Cenaze Duası" },
  { ar: "اللَّهُمَّ اغْفِرْ لِحَيِّنَا وَمَيِّتِنَا وَشَاهِدِنَا وَغَائِبِنَا وَصَغِيرِنَا وَكَبِيرِنَا وَذَكَرِنَا وَأُنْثَانَا", okunus: "Allâhümme'ğfir li-hayyinâ ve meyyitinâ ve şâhidinâ ve ğâibinâ ve sağîrinâ ve kebîrinâ ve zekerinâ ve ünsânâ.", meal: "Allah'ım! Dirimizi, ölümüzü, hazır bulunanımızı, gâibimizi, küçüğümüzü, büyüğümüzü, erkeğimizi ve kadınımızı bağışla." },
  { ar: "اللَّهُمَّ مَنْ أَحْيَيْتَهَا مِنَّا فَأَحْيِهَا عَلَى الْإِسْلَامِ وَمَنْ تَوَفَّيْتَهَا مِنَّا فَتَوَفَّهَا عَلَى الْإِيمَانِ", okunus: "Allâhümme men ahyeytehâ minnâ fe-ahyihâ ale'l-İslâmi ve men teveffeytehâ minnâ feteveffehâ ale'l-îmân.", meal: "Allah'ım! Bizden yaşattığını İslâm üzere yaşat, ecelini verdiğini de iman üzere vefat ettir." },
  { ar: "اللَّهُمَّ اغْفِرْ لَهَا وَارْحَمْهَا وَعَافِهَا وَاعْفُ عَنْهَا وَأَكْرِمْ نُزُلَهَا وَوَسِّعْ مُدْخَلَهَا وَاغْسِلْهَا بِالْمَاءِ وَالثَّلْجِ وَالْبَرَدِ وَنَقِّهَا مِنَ الْخَطَايَا كَمَا يُنَقَّى الثَّوْبُ الْأَبْيَضُ مِنَ الدَّنَسِ", okunus: "Allâhümme'ğfir lehâ verhamhâ ve âfihâ ve'fü anhâ ve ekrim nüzülehâ ve vessi' medhalehâ ve'ğsilhâ bi'l-mâi ve's-selci ve'l-bered ve nekkihâ mine'l-hatâyâ kemâ yünakkâ's-sevbü'l-ebyadu mine'd-denes.", meal: "Allah'ım! Onu bağışla, ona merhamet et, ona afiyet ver, onu affet, konduğu yeri şereflendir, kabrini genişlet; onu su, kar ve dolu ile yıka, beyaz elbisenin kirden temizlendiği gibi onu hatalardan temizle." },
  { ar: "وَأَبْدِلْهَا دَارًا خَيْرًا مِنْ دَارِهَا وَأَهْلًا خَيْرًا مِنْ أَهْلِهَا وَزَوْجًا خَيْرًا مِنْ زَوْجِهَا وَأَدْخِلْهَا الْجَنَّةَ وَأَعِذْهَا مِنْ عَذَابِ الْقَبْرِ وَعَذَابِ النَّارِ", okunus: "Ve ebdilhâ dâran hayran min dârihâ ve ehlen hayran min ehlihâ ve zevcen hayran min zevcihâ ve edhılhâ'l-cennete ve a'izhâ min azâbi'l-kabri ve azâbi'n-nâr.", meal: "Ona, bulunduğu evden daha hayırlı bir ev, daha hayırlı aile ve eş ver; onu cennete koy ve kabir ile cehennem azabından koru." },
  { section: "Kabir Sorularında Sabit Kılma Duası" },
  { ar: "اللَّهُمَّ ثَبِّتْهَا بِالْقَوْلِ الثَّابِتِ", okunus: "Allâhümme sebbithâ bi'l-kavli's-sâbit.", meal: "Allah'ım! Onu, dünyada da âhirette de sabit söz (kelime-i şehâdet) üzere sabit kıl." },
  { ar: "اللَّهُمَّ نَوِّرْ قَبْرَهَا وَوَسِّعْ مَدْخَلَهَا وَآمِنْهَا مِنْ فَزَعِ الْيَوْمِ الْأَكْبَرِ", okunus: "Allâhümmе nûvvir kabrahâ ve vessi' medhalehâ ve âminhâ min fezei'l-yevmi'l-ekber.", meal: "Allah'ım! Kabrini nurlandır, girişini genişlet ve onu büyük kıyamet gününün dehşetinden emin kıl." },
  { section: "Hatim Programı Kapsamında Kapanış" },
  { meal: "Yâ Rab! Okuduğumuz bu hatmin, yaptığımız bu duanın ve cenaze namazını kılanların dualarının sevabını, başta bu vefat eden kardeşimiz olmak üzere bütün mü'min erkek ve kadınların rûhlarına hediye eyledik; kabul eyle. Yâ Rab! Ahiret yolculuğuna çıkan bu kulunu sen bağışla, ona rahmet eyle. Onu sabit söz üzere sabit kıl, kabrini cennet bahçelerinden bir bahçe eyle. Ardında kalan ailesine, yakınlarına ve cenazesinde dua edenlere sabır, tesellî ve ecr-i azîm nasip eyle. Bizleri de onu hayırla yâd edenlerden, ölümü tefekkür edip kendini hazırlayanlardan eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const cenazeCocuk = duaDoc("CENAZE DUASI — ÇOCUK", "(Vefat etmiş çocuğun cenaze namazında okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "إِنَّا لِلَّهِ وَإِنَّا إِلَيْهِ رَاجِعُونَ", okunus: "İnnâ lillâhi ve innâ ileyhi râciûn.", meal: "Biz Allah'a aitiz ve şüphesiz O'na döneceğiz. (Bakara 156)" },
  { ar: "اللَّهُمَّ صَلِّ عَلَى مُحَمَّدٍ وَعَلَى آلِ مُحَمَّدٍ", okunus: "Allâhümme salli alâ Muhammedin ve alâ âli Muhammed.", meal: "Allah'ım! Muhammed'e ve Muhammed'in âline rahmet eyle." },
  { section: "Çocuk İçin Cenaze Duası" },
  { ar: "اللَّهُمَّ اجْعَلْهُ لَنَا فَرَطًا وَسَلَفًا وَزُخْرًا وَأَجْرًا", okunus: "Allâhümmec'alhû lenâ feratan ve selefen ve zuhran ve ecran.", meal: "Allah'ım! Onu bizim için (ahirette) önden giden, bize şefaat eden, azık ve ecir (sevap) eyle." },
  { ar: "اللَّهُمَّ اجْعَلْهُ شَافِعًا مُشَفَّعًا", okunus: "Allâhümmec'alhû şâfi'an müşeffe'an.", meal: "Allah'ım! Onu şefaat eden ve şefaati kabul edilen bir hayırlı evlât eyle." },
  { ar: "اللَّهُمَّ اغْفِرْ لَهُ وَارْحَمْهُ وَعَافِهِ وَاعْفُ عَنْهُ وَأَكْرِمْ نُزُلَهُ وَوَسِّعْ مُدْخَلَهُ", okunus: "Allâhümme'ğfir lehû verhamhû ve âfihî ve'fü anhû ve ekrim nüzülehû ve vessi' medhalehû.", meal: "Allah'ım! Onu bağışla, ona merhamet et, ona afiyet ver, onu affet, konduğu yeri şereflendir, kabrini genişlet." },
  { ar: "اللَّهُمَّ اجْعَلْهُ لِأَبَوَيْهِ أَجْرًا وَفَرَطًا وَشَفِيعًا مُجَابًا", okunus: "Allâhümme'c'alhû li-ebaveyhi ecran ve feratan ve şefîan mücâbân.", meal: "Allah'ım! Onu anne ve babası için ecir, önden giden ve duası kabul edilen şefaatçi kıl." },
  { section: "Anne-Babaya Dua" },
  { ar: "اللَّهُمَّ اصْبِرْ أَبَوَيْهِ وَأَعْقِبْهُمَا ثَوَابًا", okunus: "Allâhümme'sbir ebeveyhi ve a'kıbhümâ sevâbâ.", meal: "Allah'ım! Anne ve babasına sabır ver, onlara bu acının karşılığında sevap ve güzel bir ecir ihsan eyle." },
  { ar: "اللَّهُمَّ ارْزُقْهُمَا ذُرِّيَّةً طَيِّبَةً وَاجْبُرْ قُلُوبَهُمَا وَاشْفِ جِرَاحَهُمَا", okunus: "Allâhümme'rzukhümâ zürriyyeten tayyibeten vecbür kulûbehümâ veşfi cirâhehümâ.", meal: "Allah'ım! Onlara hayırlı ve temiz bir zürriyet ver, kalplerini tesellî et, yaralarını sar." },
  { section: "Hatim Programı Kapsamında Kapanış" },
  { meal: "Yâ Rab! Okuduğumuz bu hatmin ve yaptığımız bu duanın sevabını, bu masum yavrunun rûhuna ve bütün mü'min çocukların rûhlarına hediye eyledik; kabul eyle. Yâ Rab! Bu günahsız yavruyu ahirette anne ve babası için bir şefaatçi eyle; onu cennetinde âlemlere göster. Ailesine büyük bir sabır ve tesellî ver; bu acının her ânını ecir ve sevaba çevir. Kalbimizi senin takdirine rızayla doldur, sabredenlerle beraber eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const kabirZiyareti = duaDoc("KABİR ZİYARETİ DUASI", "(Mezarlığa gidilince okunur)", [
  { section: "Selâm Verme" },
  { ar: "السَّلَامُ عَلَيْكُمْ أَهْلَ الدِّيَارِ مِنَ الْمُؤْمِنِينَ وَالْمُسْلِمِينَ وَإِنَّا إِنْ شَاءَ اللَّهُ بِكُمْ لَاحِقُونَ", okunus: "Es-selâmü aleyküm ehle'd-diyâri mine'l-mü'minîne ve'l-müslimîn. Ve innâ inşâallâhü biküm lâhıkûn.", meal: "Ey mü'min ve müslümanlar diyarının sakinleri! Size selâm olsun. İnşâallah biz de size katılacağız." },
  { ar: "السَّلَامُ عَلَيْكُمْ دَارَ قَوْمٍ مُؤْمِنِينَ وَأَتَاكُمْ مَا تُوعَدُونَ غَدًا مُؤَجَّلُونَ وَإِنَّا إِنْ شَاءَ اللَّهُ بِكُمْ لَاحِقُونَ", okunus: "Es-selâmü aleyküm dâra kavmin mü'minîn ve etâküm mâ tû'adûne ğaden müeccelûn ve innâ inşâallâhü biküm lâhıkûn.", meal: "Selâm olsun size, ey mü'minler topluluğu! Vaat olunduğunuz şey yarın size gelecektir, er veya geç biz de inşâallah size katılacağız." },
  { section: "Dua" },
  { ar: "اللَّهُمَّ اغْفِرْ لِأَهْلِ الْقُبُورِ", okunus: "Allâhümmağfir li ehli'l-kubûr.", meal: "Allah'ım! Kabir ehlini bağışla." },
  { ar: "اللَّهُمَّ اغْفِرْ لِلْمُؤْمِنِينَ وَالْمُؤْمِنَاتِ وَالْمُسْلِمِينَ وَالْمُسْلِمَاتِ الْأَحْيَاءِ مِنْهُمْ وَالْأَمْوَاتِ", okunus: "Allâhümme'ğfir li'l-mü'minîne ve'l-mü'minât ve'l-müslimîne ve'l-müslimât el-ahyâi minhüm ve'l-emvât.", meal: "Allah'ım! Mü'min erkekleri ve mü'min kadınları, müslüman erkekleri ve müslüman kadınları, dirilerini de ölülerini de bağışla." },
  { section: "Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Şu kabristanda yatan bütün mü'min kardeşlerimize rahmet eyle. Kabirlerini nurlandır, genişlet; makamlarını yücelt. Bildiğimiz ve bilmediğimiz bütün ölmüşlerimizin rûhlarını şâd eyle. Onlara okuduğumuz Kur'ân ve yaptığımız duaların sevabını ulaştır. Bizi de onlara hayırlı bir halef kıl. Âmîn." },
]);

const kabirDua = duaDoc("KABİR ZİYARETİNDE YAPILACAK DUA", "(Mezar başında okunacak sûre ve dualar; hatim programı kapsamında okunan âyetlerin sevabı da burada ölülere hediye edilir)", [
  { section: "Ziyaret âdâbı" },
  { meal: "Kabir ziyaretine abdestli gidilir, kabirlere sırtımızı dönmeden selâm verilir, kıbleye yönelerek Kur'ân okunur ve dua edilir. Ziyaretçi, ölüyü değil; ölümü ve ahireti hatırlar, ibret alır. Peygamberimiz (s.a.v.) kabir ziyaretini unutulmaz bir ibret ve tefekkür vesilesi olarak tavsiye buyurmuştur." },
  { section: "Ziyaret Selâmı" },
  { ar: "السَّلَامُ عَلَيْكُمْ أَهْلَ الدِّيَارِ مِنَ الْمُؤْمِنِينَ وَالْمُسْلِمِينَ وَإِنَّا إِنْ شَاءَ اللَّهُ بِكُمْ لَاحِقُونَ", okunus: "Es-selâmü aleyküm ehlе'd-diyâri mine'l-mü'minîne ve'l-müslimîn, ve innâ in şâallâhü biküm lâhıkûn.", meal: "Ey mü'min ve müslümanlar diyarının sakinleri! Selâm sizlerin üzerine olsun. İnşallah biz de size katılacağız." },
  { section: "Kısa süreler" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ قُلْ هُوَ اللَّهُ أَحَدٌ * اللَّهُ الصَّمَدُ * لَمْ يَلِدْ وَلَمْ يُولَدْ * وَلَمْ يَكُنْ لَهُ كُفُوًا أَحَدٌ", okunus: "Kul hüvallâhü ehad. Allâhü's-samed. Lem yelid ve lem yûled. Ve lem yekün lehû küfüven ehad.", meal: "İhlâs Sûresi: De ki: O Allah birdir. Allah sameddir (her şey O'na muhtaçtır, O hiçbir şeye muhtaç değildir). O doğurmamıştır, doğurulmamıştır. Hiçbir şey O'na denk değildir." },
  { ar: "قُلْ أَعُوذُ بِرَبِّ الْفَلَقِ مِنْ شَرِّ مَا خَلَقَ وَمِنْ شَرِّ غَاسِقٍ إِذَا وَقَبَ وَمِنْ شَرِّ النَّفَّاثَاتِ فِي الْعُقَدِ وَمِنْ شَرِّ حَاسِدٍ إِذَا حَسَدَ", okunus: "Kul eûzü bi-rabbi'l-felak... (Felak Sûresi)", meal: "Felak Sûresi: Sabahın Rabbine sığınırım: yarattığı şeylerin şerrinden, karanlığı çöktüğü zaman gecenin şerrinden, düğümlere üfleyenlerin şerrinden ve haset ettiği zaman hasetçinin şerrinden." },
  { ar: "قُلْ أَعُوذُ بِرَبِّ النَّاسِ مَلِكِ النَّاسِ إِلَهِ النَّاسِ مِنْ شَرِّ الْوَسْوَاسِ الْخَنَّاسِ", okunus: "Kul eûzü bi-rabbi'n-nâs... (Nâs Sûresi)", meal: "Nâs Sûresi: De ki: İnsanların Rabbine, insanların Melik'ine (hükümdarına), insanların İlâhına; o sinsi vesvesecinin şerrinden sığınırım." },
  { ar: "اللَّهُمَّ صَلِّ عَلَى مُحَمَّدٍ وَعَلَى آلِ مُحَمَّدٍ كَمَا صَلَّيْتَ عَلَى إِبْرَاهِيمَ وَعَلَى آلِ إِبْرَاهِيمَ إِنَّكَ حَمِيدٌ مَجِيدٌ", okunus: "Allâhümme salli alâ Muhammedin ve alâ âli Muhammed kemâ salleyte alâ İbrâhîme ve alâ âli İbrâhîm. İnneke hamîdün mecîd.", meal: "Allah'ım! İbrâhîm'e ve İbrâhîm'in âline rahmet ettiğin gibi Muhammed'e ve Muhammed'in âline de rahmet eyle. Şüphesiz sen övülmeye lâyık, şanı yüce olansın." },
  { section: "Ölüye dua" },
  { ar: "اللَّهُمَّ اغْفِرْ لِفُلَانٍ وَارْحَمْهُ وَعَافِهِ وَاعْفُ عَنْهُ وَوَسِّعْ مُدْخَلَهُ وَأَكْرِمْ نُزُلَهُ", okunus: "Allâhümme'ğfir li-fülânin verhamhû ve âfihî ve'fü anhû ve vessi' medhalehû ve ekrim nüzülehû.", meal: "Allah'ım! (Burada yatan) kulunu bağışla, ona merhamet et, ona afiyet ver, onu affet, kabrini genişlet, konduğu yeri şereflendir." },
  { ar: "رَبِّ اغْفِرْ لِي وَلِوَالِدَيَّ وَلِلْمُؤْمِنِينَ يَوْمَ يَقُومُ الْحِسَابُ", okunus: "Rabbi'ğfir lî ve li-vâlideyye ve li'l-mü'minîne yevme yekûmü'l-hisâb.", meal: "Rabbim! Hesabın görüleceği gün beni, anne-babamı ve bütün mü'minleri bağışla. (İbrâhîm 41)" },
  { ar: "اللَّهُمَّ اغْفِرْ لِلْمُؤْمِنِينَ وَالْمُؤْمِنَاتِ وَالْمُسْلِمِينَ وَالْمُسْلِمَاتِ الْأَحْيَاءِ مِنْهُمْ وَالْأَمْوَاتِ", okunus: "Allâhümme'ğfir li'l-mü'minîne ve'l-mü'minât ve'l-müslimîne ve'l-müslimât, el-ahyâi minhüm ve'l-emvât.", meal: "Allah'ım! Bütün mü'min erkekleri, mü'min kadınları, müslüman erkekleri ve müslüman kadınları; onlardan hayatta olanları da vefat etmiş olanları da bağışla." },
  { section: "Hatim Programı Kapsamında Kapanış" },
  { meal: "Yâ Rab! Bu kabirlerde yatan ve bütün mü'min mezarlarında bulunan kardeşlerimize, okuduğumuz bu hatmin âyetlerinin ve yaptığımız bu duaların sevabını hediye eyledik; kabul eyle. Yâ Rab! Onların kabirlerini nurlandır, genişlet; onları kabir azabından ve cehennem azabından muhafaza eyle. Bizleri de onların hayırlı amelleriyle anan, ölülerini hayırla yâd eden kullarından eyle. Bizleri onlara kavuşturacak ölümü, bize rahmetine kavuşma vesilesi kıl. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const telkin = duaDoc("TELKİN DUASI", "(Definden sonra mezar başında okunur; hatim programı kapsamında vefat edenin kabri başında da okunur)", [
  { section: "Telkine başlarken" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "اللَّهُمَّ صَلِّ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Allâhümme salli alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Allah'ım! Efendimiz Muhammed'e, âline ve ashâbının tümüne rahmet eyle." },
  { section: "Telkin Metni" },
  { ar: "يَا عَبْدَ اللَّهِ يَا ابْنَ أَمَةِ اللَّهِ، اذْكُرْ مَا خَرَجْتَ عَلَيْهِ مِنَ الدُّنْيَا: شَهَادَةَ أَنْ لَا إِلَهَ إِلَّا اللَّهُ وَأَنَّ مُحَمَّدًا رَسُولُ اللَّهِ", okunus: "Yâ abdallâhi yâ ibne emetillâh! Uzkur mâ haracte aleyhi mine'd-dünyâ: şehâdete en lâ ilâhe illallâhü ve enne Muhammeden Resûlullâh.", meal: "Ey Allah'ın kulu, Allah'ın cariyesinin oğlu! Dünyadan ayrılırken üzerinde bulunduğun (inanç) şu idi: Allah'tan başka ilâh olmadığına ve Muhammed'in (s.a.v.) Allah'ın Resûlü olduğuna şehâdet etmek." },
  { ar: "وَأَنَّ الْجَنَّةَ حَقٌّ وَالنَّارَ حَقٌّ وَالْبَعْثَ حَقٌّ وَأَنَّ السَّاعَةَ آتِيَةٌ لَا رَيْبَ فِيهَا وَأَنَّ اللَّهَ يَبْعَثُ مَنْ فِي الْقُبُورِ", okunus: "Ve enne'l-cennete hakkun ve'n-nâre hakkun ve'l-ba'se hakkun ve enne's-sâ'ate âtiyetün lâ raybe fîhâ ve ennallâhe yeb'asü men fi'l-kubûr.", meal: "Ve cennet haktır, cehennem haktır, yeniden diriliş haktır; kıyamet mutlaka gelecektir, onda şüphe yoktur ve Allah kabirlerdeki herkesi diriltecektir." },
  { ar: "رَضِيتُ بِاللَّهِ رَبًّا وَبِالْإِسْلَامِ دِينًا وَبِمُحَمَّدٍ نَبِيًّا وَرَسُولًا", okunus: "Radîtü billâhi Rabben ve bi'l-İslâmi dînen ve bi-Muhammedin sallallâhü aleyhi ve sellem nebiyyen ve resûlen.", meal: "Rabb olarak Allah'a, din olarak İslâm'a, peygamber ve resûl olarak Muhammed'e (s.a.v.) razı oldum." },
  { ar: "يَا عَبْدَ اللَّهِ، رَبُّكَ اللَّهُ، دِينُكَ الْإِسْلَامُ، نَبِيُّكَ مُحَمَّدٌ، قِبْلَتُكَ الْكَعْبَةُ، كِتَابُكَ الْقُرْآنُ", okunus: "Yâ abdallâh, rabbüke'llâh, dînüke'l-İslâm, nebiyyüke Muhammed, kıbletüke'l-Ka'betü, kitâbüke'l-Kur'ân.", meal: "Ey Allah'ın kulu! Rabbin Allah'tır, dinin İslâm'dır, peygamberin Muhammed'dir, kıblen Kâbe'dir, kitabın Kur'ân'dır." },
  { section: "Ölüye Dua" },
  { ar: "اللَّهُمَّ اغْفِرْ لَهُ وَارْحَمْهُ وَثَبِّتْهُ بِالْقَوْلِ الثَّابِتِ", okunus: "Allâhümme'ğfir lehû verhamhû ve sebbithü bi'l-kavli's-sâbit.", meal: "Allah'ım! Onu bağışla, ona merhamet et, onu sâbit sözle (kelime-i şehâdet ile) sabit kıl." },
  { ar: "اللَّهُمَّ نَوِّرْ قَبْرَهُ وَوَسِّعْ مَدْخَلَهُ وَآمِنْهُ مِنْ عَذَابِ الْقَبْرِ وَفَزَعِ الْيَوْمِ الْأَكْبَرِ", okunus: "Allâhümme nûvvir kabrahu ve vessi' medhalehû ve âminhü min azâbi'l-kabri ve fezei'l-yevmi'l-ekber.", meal: "Allah'ım! Kabrini nurlandır, girişini genişlet ve onu kabir azabından ve büyük kıyamet gününün dehşetinden emin kıl." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe)" },
  { meal: "Yâ Rab! Bu kulunun kabrini nurlandır, Münker ve Nekir suallerinde ona yardım eyle, imanını muhafaza buyur, kabir azabından muhafaza eyle. Ona sabrı ve imanı şefaatçi kıl. Okuduğumuz hatmin âyetlerini, telkini ve duaların sevabını rûhuna hediye eyledik; kabul eyle. Yâ Rab! Onu, dostlarınla birlikte huzuruna kabul eyle; kabrini cennet bahçelerinden bir bahçe, dünya evinden daha hayırlı bir durak eyle. Geride kalan yakınlarına sabır, bağışlananlardan olmayı ve cennetine girmeyi nasip eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

function sureDoc(titleText, arabic, meal) {
  return [
    title(titleText),
    blank(),
    ...(Array.isArray(arabic) ? arabic : [arabic]).map((a) => p(a, { rtl: true, after: 200 })),
    blank(),
    section("MEÂLİ"),
    ...(Array.isArray(meal) ? meal : [meal]).map((m) => p(m, { italics: true, color: "44564e" })),
  ];
}

const nisanDuasi = duaDoc("NİŞAN DUASI", "(Nişan töreninde okunur; hatim programı kapsamında evlilik hayırı için de okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "الْحَمْدُ لِلَّهِ رَبِّ الْعَالَمِينَ وَالصَّلَاةُ وَالسَّلَامُ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Elhamdülillâhi rabbi'l-âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Hamd, âlemlerin Rabbi Allah'a mahsustur. Salât ve selâm, efendimiz Muhammed'in, âlinin ve ashâbının üzerine olsun." },
  { section: "Âyet (Rûm 21)" },
  { ar: "وَمِنْ آيَاتِهِ أَنْ خَلَقَ لَكُمْ مِنْ أَنْفُسِكُمْ أَزْوَاجًا لِتَسْكُنُوا إِلَيْهَا وَجَعَلَ بَيْنَكُمْ مَوَدَّةً وَرَحْمَةً إِنَّ فِي ذَلِكَ لَآيَاتٍ لِقَوْمٍ يَتَفَكَّرُونَ", okunus: "Ve min âyâtihî en haleka leküm min enfüsiküm ezvâcen li-teskünû ileyhâ ve ceale beyneküm meveddeten ve rahmeh. İnne fî zâlike le-âyâtin li-kavmin yetefekkerûn.", meal: "Kaynaşıp huzur bulasınız diye sizin için kendi türünüzden eşler yaratıp aranıza sevgi ve merhamet koyması da O'nun âyetlerindendir. Şüphesiz bunda düşünen bir toplum için ibretler vardır." },
  { section: "Birliktelik ve Aile Duası" },
  { ar: "رَبَّنَا هَبْ لَنَا مِنْ أَزْوَاجِنَا وَذُرِّيَّاتِنَا قُرَّةَ أَعْيُنٍ وَاجْعَلْنَا لِلْمُتَّقِينَ إِمَامًا", okunus: "Rabbenâ heb lenâ min ezvâcinâ ve zürriyyâtinâ kurrete a'yünin vec'alnâ li'l-müttakîne imâmâ.", meal: "Rabbimiz! Bize eşlerimizden ve çocuklarımızdan göz aydınlığı olacak kimseler ihsan et ve bizi takva sahiplerine önder eyle. (Furkân 74)" },
  { section: "Dua" },
  { ar: "بَارَكَ اللَّهُ لَكَ وَبَارَكَ عَلَيْكَ وَجَمَعَ بَيْنَكُمَا فِي خَيْرٍ", okunus: "Bârekallâhü leke ve bârake aleyke ve cemea beynekümâ fî hayr.", meal: "Allah sana bereket versin, senin üzerine bereket indirsin ve ikinizi hayır üzere birleştirsin. (Tirmizî, Nikâh 35)" },
  { ar: "اللَّهُمَّ أَلِّفْ بَيْنَ قُلُوبِهِمَا وَأَصْلِحْ بَيْنَهُنَّ وَاهْدِهِمْ سُبُلَ السَّلَامِ", okunus: "Allâhümme ellif beyne kulûbihimâ ve aslih beynenâ ve-hdihim sübüle's-selâm.", meal: "Allah'ım! İkisinin kalplerini birbirine ısındır, aralarını ıslah et ve onları esenlik yollarına hidayet eyle." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Bu nişanlanan gençlerin birbirine olan sevgisini artır, aralarına muhabbet ve merhamet nasip eyle. Onları hayırlı, helâl ve mutlu bir evliliğe erdir. Ailelerini ve bu mecliste bulunanları rızanla sevindir. Bu nişanı, inşâallah hayırlı bir nikâha ve huzurlu bir yuva kurulmasına vesile eyle. Okuduğumuz hatmin ve bu duanın sevabını bütün evlilik hayırlarına aday gençlere ve âilelerine hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const nikahDuasi = duaDoc("NİKÂH DUASI — ARAPÇA VE TÜRKÇE", "(Nikâh akdi sonrasında okunur; hatim programı kapsamında evlilik hayırı için de okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ الْحَمْدُ لِلَّهِ وَالصَّلَاةُ وَالسَّلَامُ عَلَى رَسُولِ اللَّهِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm. Elhamdü lillâhi ves-salâtü ves-selâmü alâ Resûlillâh.", meal: "Rahman ve Rahîm olan Allah'ın adıyla. Hamd Allah'a mahsustur; salât ve selâm Allah'ın Resûlü'nün üzerine olsun." },
  { ar: "اللَّهُمَّ صَلِّ عَلَى مُحَمَّدٍ وَعَلَى آلِ مُحَمَّدٍ كَمَا صَلَّيْتَ عَلَى إِبْرَاهِيمَ وَعَلَى آلِ إِبْرَاهِيمَ إِنَّكَ حَمِيدٌ مَجِيدٌ", okunus: "Allâhümme salli alâ Muhammedin ve alâ âli Muhammed kemâ salleyte alâ İbrâhîme ve alâ âli İbrâhîm. İnneke hamîdün mecîd.", meal: "Allah'ım! İbrâhîm'e ve İbrâhîm'in âline rahmet ettiğin gibi Muhammed'e ve Muhammed'in âline de rahmet eyle. Şüphesiz sen övülmeye lâyık, şanı yüce olansın." },
  { section: "Âyet (Rûm 21)" },
  { ar: "وَمِنْ آيَاتِهِ أَنْ خَلَقَ لَكُمْ مِنْ أَنْفُسِكُمْ أَزْوَاجًا لِتَسْكُنُوا إِلَيْهَا وَجَعَلَ بَيْنَكُمْ مَوَدَّةً وَرَحْمَةً إِنَّ فِي ذَلِكَ لَآيَاتٍ لِقَوْمٍ يَتَفَكَّرُونَ", okunus: "Ve min âyâtihî en haleka leküm min enfüsiküm ezvâcen li-teskünû ileyhâ ve ceale beyneküm meveddeten ve rahmeh. İnne fî zâlike le-âyâtin li-kavmin yetefekkerûn.", meal: "Kaynaşıp huzur bulasınız diye sizin için kendi türünüzden eşler yaratıp aranıza sevgi ve merhamet koyması da O'nun âyetlerindendir. Şüphesiz bunda düşünen bir toplum için ibretler vardır." },
  { section: "Aile ve Zürriyet Duası" },
  { ar: "رَبَّنَا هَبْ لَنَا مِنْ أَزْوَاجِنَا وَذُرِّيَّاتِنَا قُرَّةَ أَعْيُنٍ وَاجْعَلْنَا لِلْمُتَّقِينَ إِمَامًا", okunus: "Rabbenâ heb lenâ min ezvâcinâ ve zürriyyâtinâ kurrete a'yünin vec'alnâ li'l-müttakîne imâmâ.", meal: "Rabbimiz! Bize eşlerimizden ve çocuklarımızdan göz aydınlığı olacak kimseler ihsan et ve bizi takva sahiplerine önder eyle. (Furkân 74)" },
  { section: "Nikâh Duası (Hadis-i Şerif'ten)" },
  { ar: "بَارَكَ اللَّهُ لَكَ وَبَارَكَ عَلَيْكَ وَجَمَعَ بَيْنَكُمَا فِي خَيْرٍ", okunus: "Bârekallâhü leke ve bârake aleyke ve cemea beynekümâ fî hayr.", meal: "Allah sana bereket versin, üzerine bereket indirsin ve ikinizi hayır üzere birleştirsin. (Tirmizî, Nikâh 35)" },
  { ar: "اللَّهُمَّ اجْعَلْ بَيْنَهُمَا مَوَدَّةً وَرَحْمَةً وَأَلِّفْ بَيْنَ قُلُوبِهِمَا", okunus: "Allâhümme'c'al beynehümâ meveddeten ve rahmeh ve ellif beyne kulûbihimâ.", meal: "Allah'ım! İkisinin arasında sevgi ve merhamet eyle, kalplerini birbirine ısındır." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Bu aziz nikâhla bir araya gelen kullarını karşılıklı sevgi, saygı ve merhametle donat. Aralarına muhabbet ve huzur nasip eyle. Onları sağlıklı, sâlih evlâtlarla rızıklandır. Bu yuvayı cennet bahçelerinden bir bahçe eyle. Aileleri, sıla-i rahmi ve tüm müslümanları bu hayırlı evlilikte sevindir. Geçimlerinde bereket, kalplerinde karar ve mutluluk ver. Okuduğumuz hatmin sevabını bu yuvanın, gençlerin ve bütün evlilik hayırlarının üzerine hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const gelinUgurlama = duaDoc("GELİN UĞURLAMA DUASI", "(Gelinin baba evinden uğurlanırken okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "اللَّهُمَّ صَلِّ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Allâhümme salli alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Allah'ım! Efendimiz Muhammed'e, âline ve ashâbının tümüne rahmet eyle." },
  { section: "Âyet (Furkân 74)" },
  { ar: "رَبَّنَا هَبْ لَنَا مِنْ أَزْوَاجِنَا وَذُرِّيَّاتِنَا قُرَّةَ أَعْيُنٍ وَاجْعَلْنَا لِلْمُتَّقِينَ إِمَامًا", okunus: "Rabbenâ heb lenâ min ezvâcinâ ve zürriyyâtinâ kurrete a'yünin vec'alnâ li'l-müttakîne imâmâ.", meal: "Rabbimiz! Bize eşlerimizden ve çocuklarımızdan göz aydınlığı olacak kimseler ihsan et ve bizi takva sahiplerine önder eyle." },
  { section: "Dua" },
  { ar: "بَارَكَ اللَّهُ لَكُمَا وَبَارَكَ عَلَيْكُمَا وَجَمَعَ بَيْنَكُمَا فِي خَيْرٍ", okunus: "Bârekallâhü lekümâ ve bârake aleykümâ ve cemea beynekümâ fî hayr.", meal: "Allah ikinize bereket versin, üzerinize bereket indirsin ve ikinizi hayır üzere birleştirsin." },
  { ar: "اللَّهُمَّ اجْعَلْ بَيْنَهُمَا مَوَدَّةً وَرَحْمَةً وَسَكِينَةً", okunus: "Allâhümme'c'al beynehümâ meveddeten ve rahmeten ve sekîneh.", meal: "Allah'ım! Aralarında sevgi, merhamet ve huzur kıl." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Evinden, yurdundan ayrılan bu gelin kızımıza hayırlı yolculuklar ve hayırlı bir yuva nasip eyle. Onu, gittiği yerde güler yüzle, huzurla ve bereketle karşıla. Anne-babasının, büyüklerinin hayır duasını üzerinden eksik eyleme. Bu hayırlı birlikteliği güzelliklerle süslendir; iki aileyi muhabbetle kaynaştır. Gelinimize de güveyimize de sıhhat ve afiyetle, mutlu ve huzurlu bir ömürle rızıklandır. Okuduğumuz hatmin ve bu duanın sevabını bu hayırlı yuvaya ve bütün evlilik hayırlarına hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const yeniDogan = duaDoc("YENİ DOĞAN ÇOCUK DUASI", "(Doğum sonrasında kulağına ezan ve dua okunur; hatim programı kapsamında çocuklar için de okunur)", [
  { section: "Doğum âdâbı" },
  { meal: "Doğan çocuğun sağ kulağına ezan, sol kulağına kâmet okunması sünnettir. Çocuk için hayır dua edilir, yedinci gününde isim verilir ve akîka kurbanı kesilir." },
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "الْحَمْدُ لِلَّهِ رَبِّ الْعَالَمِينَ وَالصَّلَاةُ وَالسَّلَامُ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Elhamdülillâhi rabbi'l-âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Hamd, âlemlerin Rabbi Allah'a mahsustur. Salât ve selâm, efendimiz Muhammed'in, âlinin ve ashâbının üzerine olsun." },
  { section: "Âyet (Âl-i İmrân 36)" },
  { ar: "وَالَّذِي وَلَدْتُ أُعِيذُهَا بِكَ وَذُرِّيَّتَهَا مِنَ الشَّيْطَانِ الرَّجِيمِ", okunus: "Ve'l-lezî veledtü uîzühâ bike ve zürriyyetehâ mine'ş-şeytâni'r-racîm.", meal: "Doğurduğum çocuğu ve onun soyunu kovulmuş şeytanın şerrinden sana sığındırırım." },
  { section: "Yeni Doğan İçin Dua" },
  { ar: "اللَّهُمَّ اجْعَلْهُ مُبَارَكًا وَاجْعَلْهُ صَالِحًا مُؤْمِنًا", okunus: "Allâhümme'c'alhû mübareken ve'c'alhû sâlihan mü'minâ.", meal: "Allah'ım! Onu bereketli, sâlih ve mü'min bir kul eyle." },
  { ar: "أُعِيذُهُ بِاللَّهِ الْوَاحِدِ مِنْ شَرِّ كُلِّ حَاسِدٍ", okunus: "Uîzühû billâhi'l-vâhidi min şerri külli hâsidin.", meal: "Onu, bir olan Allah'a sığındırırım; her hasetçinin şerrinden." },
  { ar: "اللَّهُمَّ أَنْبِتْهُ نَبَاتًا حَسَنًا وَاجْعَلْهُ قُرَّةَ عَيْنٍ لِوَالِدَيْهِ", okunus: "Allâhümme enbit-hü nebâten hasenen ve'c'alhü kurrete aynin li-vâlideyh.", meal: "Allah'ım! Onu güzel bir bitki gibi (güzelce) büyüt ve onu anne-babasının göz aydınlığı eyle." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Yeni doğan bu yavruyu sağlıklı, sâlih, hayırlı ve bahtı açık bir kul eyle. Anne ve babasına hayırlı bir evlât kıl. Onu, âilesine ve ümmetine faydalı eyle. Gönlünü imanla, ahlâkını güzellikle doldur. Onu her türlü hastalık, nazar ve kötülükten muhafaza eyle. Okuduğumuz hatmin ve bu duanın sevabını bütün çocuklarımıza, gençlerimize ve onların hayırlı yetişmelerine hediye eyledik; kabul eyle. Yâ Rab, bu masum yavruyu ve anne-babasını rahmetinle kuşat. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const sunnetDuasi = duaDoc("ÇOCUK SÜNNET DUASI", "(Sünnet merasiminde okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "اللَّهُمَّ صَلِّ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Allâhümme salli alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Allah'ım! Efendimiz Muhammed'e, âline ve ashâbının tümüne rahmet eyle." },
  { section: "Dua" },
  { ar: "اللَّهُمَّ اجْعَلْهُ صَحِيحًا سَلِيمًا مُعَافًى وَاجْعَلْهُ مِنَ الصَّالِحِينَ", okunus: "Allâhümme'c'alhû sahîhan selîmen muâfâ ve'c'alhû mine's-sâlihîn.", meal: "Allah'ım! Onu sağlıklı, sâlim ve âfiyette kıl; onu sâlihlerden eyle." },
  { ar: "اللَّهُمَّ بَارِكْ فِيهِ وَاجْعَلْهُ مِنَ الذُّرِّيَّةِ الطَّيِّبَةِ", okunus: "Allâhümme bârik fîhi ve'c'alhû mine'z-zürriyyeti't-tayyibeh.", meal: "Allah'ım! Onu bereketlendir ve onu temiz (hayırlı) nesilden eyle." },
  { ar: "اللَّهُمَّ احْفَظْهُ وَاجْعَلْهُ صَحِيحًا مُعَافًى وَأَطْلِ عَلَيْهِ الْعَافِيَةَ وَالْبَرَكَةَ", okunus: "Allâhümmehfazhû ve'c'alhû sahîhan muâfâ ve atlı aleyhi'l-âfiyete ve'l-berekeh.", meal: "Allah'ım! Onu koru, sağlıklı ve âfiyette kıl; üzerine afiyet ve bereket indir." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Sünnet olan bu yavruyu sıhhat ve âfiyetle büyüt, yarasını çabucak iyileştir. Onu Peygamber Efendimizin (s.a.v.) sünnetine bağlı, dinine ve vatanına hayırlı bir evlât eyle. Ailesine sabır ve sabrun güzeli ihsan eyle. Okuduğumuz hatmin ve bu duanın sevabını bu yavruya, onun âilesine ve bütün hayırlı yetişen çocuklara hediye eyledik; kabul eyle. Bu merasimde bulunan herkese hayır ve bereket ver. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const pazarDuasi = duaDoc("PAZAR KURULUŞ DUASI", "(Pazar ve çarşı kurulurken / esnaf için okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "الْحَمْدُ لِلَّهِ رَبِّ الْعَالَمِينَ وَالصَّلَاةُ وَالسَّلَامُ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Elhamdülillâhi rabbi'l-âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Hamd, âlemlerin Rabbi Allah'a mahsustur. Salât ve selâm, efendimiz Muhammed'in, âlinin ve ashâbının üzerine olsun." },
  { section: "Korunma Duası" },
  { ar: "بِسْمِ اللَّهِ الَّذِي لَا يَضُرُّ مَعَ اسْمِهِ شَيْءٌ فِي الْأَرْضِ وَلَا فِي السَّمَاءِ وَهُوَ السَّمِيعُ الْعَلِيمُ", okunus: "Bismillâhi'llezî lâ yedurru me'a'smihî şey'ün fi'l-erdı ve lâ fi's-semâi ve hüve's-semî'ul-alîm.", meal: "Adı anıldığında yerde ve gökte hiçbir şeyin zarar veremeyeceği Allah'ın adıyla. O, işitendir, bilendir." },
  { section: "Âyet (Nûr 36-37)" },
  { ar: "رِجَالٌ لَا تُلْهِيهِمْ تِجَارَةٌ وَلَا بَيْعٌ عَنْ ذِكْرِ اللَّهِ وَإِقَامِ الصَّلَاةِ", okunus: "Ricâlün lâ tülhîhim ticâretün ve lâ bey'un an zikrillâhi ve ikâmi's-salât.", meal: "Onlar öyle kimselerdir ki ne ticaret ne de alışveriş onları Allah'ı anmaktan, namazı kılmaktan alıkoyar. (Nûr 37)" },
  { section: "Bereket Duası" },
  { ar: "اللَّهُمَّ بَارِكْ لَنَا فِي مَدِينَتِنَا وَفِي ثِمَارِنَا وَفِي مَكِيلِنَا وَفِي مُدِّنَا وَصَاعِنَا", okunus: "Allâhümme bârik lenâ fî medînetinâ ve fî simârinâ ve fî mekîlinâ ve fî müddinâ ve sâinâ.", meal: "Allah'ım! Şehrimizi, meyvelerimizi, ölçülerimizi ve bütün rızıklarımızı bizim için bereketlendir." },
  { ar: "اللَّهُمَّ ارْزُقْنَا رِزْقًا حَلَالًا طَيِّبًا وَبَارِكْ لَنَا فِيهِ", okunus: "Allâhümme'rzuknâ rizkan halâlen tayyiben ve bârik lenâ fîh.", meal: "Allah'ım! Bize helâl ve temiz rızık ver ve onu bizim için bereketli kıl." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Bu pazar ve çarşılarda alışveriş yapan tüm esnaf ve müşterilere hayırlı ve bereketli kazançlar nasip eyle. Alın teriyle kazanılan rızıkları helâl ve temiz kıl. Ticaretin ölçüsünü ve tartısını doğrulukla yapmayı nasip eyle. Esnaflarımızı kazadan, belâdan, hırsızlıktan ve her türlü kötülükten muhafaza eyle. Okuduğumuz hatmin ve bu duanın sevabını bütün esnaf ve sanatkârlarımıza, geçimini helâl yoldan kazanan kullarına hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const isyeriAcilis = duaDoc("İŞ YERİ AÇILIŞ / TEMEL ATMA DUASI", "(İş yeri açılışında ve inşaat temelinde okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "الْحَمْدُ لِلَّهِ رَبِّ الْعَالَمِينَ وَالصَّلَاةُ وَالسَّلَامُ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Elhamdülillâhi rabbi'l-âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Hamd, âlemlerin Rabbi Allah'a mahsustur. Salât ve selâm, efendimiz Muhammed'in, âlinin ve ashâbının üzerine olsun." },
  { section: "Âyet (Tâhâ 25-26)" },
  { ar: "رَبِّ اشْرَحْ لِي صَدْرِي وَيَسِّرْ لِي أَمْرِي", okunus: "Rabbişrah lî sadrî ve yessir lî emrî.", meal: "Rabbim! Gönlüme genişlik ver, işimi kolaylaştır." },
  { section: "Âyet (Âl-i İmrân 159)" },
  { ar: "فَإِذَا عَزَمْتَ فَتَوَكَّلْ عَلَى اللَّهِ إِنَّ اللَّهَ يُحِبُّ الْمُتَوَكِّلِينَ", okunus: "Fe izâ azemte fetevekkel alellâh. İnnallâhe yühibbü'l-mütevekkilîn.", meal: "Kararını verince artık Allah'a tevekkül et. Şüphesiz Allah tevekkül edenleri sever." },
  { section: "Âyet (Talâk 2-3)" },
  { ar: "وَمَنْ يَتَّقِ اللَّهَ يَجْعَلْ لَهُ مَخْرَجًا وَيَرْزُقْهُ مِنْ حَيْثُ لَا يَحْتَسِبُ وَمَنْ يَتَوَكَّلْ عَلَى اللَّهِ فَهُوَ حَسْبُهُ", okunus: "Ve men yetteki'llâhe yec'al lehû mahracen ve yerzukhü min haysü lâ yahtesib. Ve men yetevekkel alellâhi fe-hüve hasbüh.", meal: "Kim Allah'tan sakınırsa, Allah ona bir çıkış yolu gösterir ve onu ummadığı yerden rızıklandırır. Kim Allah'a tevekkül ederse O, ona yeter." },
  { section: "Bereket ve Korunma Duası" },
  { ar: "اللَّهُمَّ بَارِكْ فِي هَذَا الْمَكَانِ وَاجْعَلْهُ مُبَارَكًا وَافْتَحْ عَلَيْهِ أَبْوَابَ الْخَيْرِ", okunus: "Allâhümme bârik fî hâzâ'l-mekân ve'c'alhü mübareken ve'fteh aleyhi ebvâbe'l-hayr.", meal: "Allah'ım! Bu mekânı bereketlendir, onu mübarek kıl ve üzerine hayır kapılarını aç." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Bu iş yerinin / binanın açılışını hayırlı ve uğurlu eyle. Sahiplerine helâl, bereketli ve devamlı kazanç nasip eyle. Burada çalışanlara hayır ve kolaylık ver. Her türlü kazadan, belâdan ve haksızlıktan muhafaza eyle. Bu işi ve bu yapıyı, sana itaat ve iyiliğe vesile kıl. Okuduğumuz hatmin ve bu duanın sevabını bu işin sahiplerine, çalışanlarına ve bütün helâl kazanç sahiplerine hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const yagmurDuasi = duaDoc("YAĞMUR (İSTİSKÂ) DUASI", "(Kuraklıkta yağmur yağması için okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "الْحَمْدُ لِلَّهِ رَبِّ الْعَالَمِينَ وَالصَّلَاةُ وَالسَّلَامُ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Elhamdülillâhi rabbi'l-âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Hamd, âlemlerin Rabbi Allah'a mahsustur. Salât ve selâm, efendimiz Muhammed'in, âlinin ve ashâbının üzerine olsun." },
  { section: "İstiskâ Duası (Hadis-i Şerif'ten)" },
  { ar: "اللَّهُمَّ اسْقِنَا الْغَيْثَ الْمُغِيثَ الْمَرِيعَ النَّافِعَ غَيْرَ الضَّارِّ عَاجِلًا غَيْرَ آجِلٍ", okunus: "Allâhümme'skınâ'l-ğayse'l-muğîse'l-merî'a'n-nâfi'a ğayra'd-dârr, âcilen ğayra âcil.", meal: "Allah'ım! Bize yardım getiren, bereketli, faydalı, zararsız, gecikmeyen yağmur yağdır." },
  { ar: "اللَّهُمَّ أَغِثْنَا، اللَّهُمَّ أَغِثْنَا، اللَّهُمَّ أَغِثْنَا", okunus: "Allâhümme eğisnâ, Allâhümme eğisnâ, Allâhümme eğisnâ.", meal: "Allah'ım! Bize yardım et (yağmur ver), üç kez." },
  { ar: "اللَّهُمَّ اسْقِنَا وَلَا تَجْعَلْنَا مِنَ الْقَانِطِينَ", okunus: "Allâhümme'skınâ ve lâ tec'alnâ mine'l-kânitîn.", meal: "Allah'ım! Bize yağmur ver; bizi (rahmetinden) ümidini kesenlerden eyleme." },
  { section: "Tevbe ve İstiğfar" },
  { ar: "اللَّهُمَّ اغْفِرْ لَنَا ذُنُوبَنَا وَتُبْ عَلَيْنَا وَارْزُقْنَا مِنْ رَحْمَتِكَ", okunus: "Allâhümme'ğfir lenâ zünûbenâ ve tüb aleynâ ve'rzuknâ min rahmetik.", meal: "Allah'ım! Günahlarımızı bağışla, tövbemizi kabul et ve bize rahmetinden rızık ver." },
  { ar: "اسْتَغْفِرُوا رَبَّكُمْ إِنَّهُ كَانَ غَفَّارًا يُرْسِلِ السَّمَاءَ عَلَيْكُمْ مِدْرَارًا", okunus: "İsteğfirû rabbeküm innehû kâne ğaffârâ. Yursili's-semâe aleyküm midrârâ.", meal: "Rabbinizden mağfiret dileyin; şüphesiz O, çok bağışlayandır. Üzerinize gökten bol bol yağmur göndersin. (Nûh 10-11)" },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Toprakları susamış olan memleketimize bereket yağmurları ihsan eyle. Kuruyan kuyuları, nehirleri, gölleri ve barajları doldur. Ekinlerimize, hayvanlarımıza ve tüm canlılara rahmetinle yağmur gönder. Bizi rahmetinden ümitsiz eyleme. Okuduğumuz hatmin ve bu duanın sevabını bütün kuraklık çeken beldelere, çiftçilerimize ve emekçilerimize hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const umreUgurlama = duaDoc("UMRE UĞURLAMA DUASI", "(Umreye gideni uğurlarken okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "الْحَمْدُ لِلَّهِ رَبِّ الْعَالَمِينَ وَالصَّلَاةُ وَالسَّلَامُ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Elhamdülillâhi rabbi'l-âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Hamd, âlemlerin Rabbi Allah'a mahsustur. Salât ve selâm, efendimiz Muhammed'in, âlinin ve ashâbının üzerine olsun." },
  { section: "Uğurlama Duası (Hadis-i Şerif'ten)" },
  { ar: "زَوَّدَكَ اللَّهُ التَّقْوَى وَغَفَرَ ذَنْبَكَ وَيَسَّرَ لَكَ الْخَيْرَ حَيْثُ كُنْتَ", okunus: "Zevvedekellâhü't-takvâ ve ğefere zenbeke ve yessera leke'l-hayra haysü mâ künt.", meal: "Allah sana takvâyı azık etsin, günahını bağışlasın ve nerede olursan ol hayrı sana kolaylaştırsın. (Tirmizî, Daavât 45)" },
  { section: "Beytullah Ziyareti Duası" },
  { ar: "اللَّهُمَّ اجْعَلْ عُمْرَتَهُ مَقْبُولَةً وَسَعْيَهُ مَشْكُورًا وَذَنْبَهُ مَغْفُورًا", okunus: "Allâhümme'c'al umretehû makbûleten ve sa'yehû meşkûran ve zenbehû mağfûran.", meal: "Allah'ım! Onun umresini makbul, sa'yini meşkûr (kabul gören) ve günahını bağışlanmış eyle." },
  { ar: "اللَّهُمَّ سَلِّمْهُ فِي ذَهَابِهِ وَإِيَابِهِ", okunus: "Allâhümme sellimhü fî zehâbihî ve iyâbih.", meal: "Allah'ım! Onu gidişinde de dönüşünde de selâmete erdir." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Umreye giden kardeşimize hayırlı yolculuklar nasip eyle. Onu emin bir şekilde Beytullah'a ulaştır. Umresini ve tavafını makbul, sa'yini meşkûr, duasını müstecab eyle. Onu sağlık ve afiyetle, iman ve sürurla yurduna döndür. Âilesine ve tüm müslümanlara hayırla dönmesini nasip eyle. Okuduğumuz hatmin ve bu duanın sevabını onun ve bütün umre yapacak müslümanların üzerine hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const haciUgurlama = duaDoc("HACI UĞURLAMA DUASI", "(Hacca gideni uğurlarken okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "الْحَمْدُ لِلَّهِ رَبِّ الْعَالَمِينَ وَالصَّلَاةُ وَالسَّلَامُ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Elhamdülillâhi rabbi'l-âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Hamd, âlemlerin Rabbi Allah'a mahsustur. Salât ve selâm, efendimiz Muhammed'in, âlinin ve ashâbının üzerine olsun." },
  { section: "Uğurlama Duası" },
  { ar: "زَوَّدَكَ اللَّهُ التَّقْوَى وَغَفَرَ ذَنْبَكَ وَيَسَّرَ لَكَ الْخَيْرَ حَيْثُ كُنْتَ", okunus: "Zevvedekellâhü't-takvâ ve ğefere zenbeke ve yessera leke'l-hayra haysü mâ künt.", meal: "Allah sana takvâyı azık etsin, günahını bağışlasın ve nerede olursan ol hayrı sana kolaylaştırsın." },
  { ar: "اللَّهُمَّ اجْعَلْ حَجًّا مَبْرُورًا وَسَعْيًا مَشْكُورًا وَذَنْبًا مَغْفُورًا", okunus: "Allâhümme'c'al haccen mebrûran ve sa'yan meşkûran ve zenben mağfûran.", meal: "Allah'ım! Haccını makbul, sa'yini meşkûr (kabul gören) ve günahını bağışlanmış eyle." },
  { ar: "اللَّهُمَّ سَلِّمْهُ إِلَى بَيْتِكَ الْحَرَامِ وَسَلِّمْهُ إِلَى أَهْلِهِ", okunus: "Allâhümme sellimhü ilâ beytike'l-harâm ve sellimhü ilâ ehlih.", meal: "Allah'ım! Onu Beyt-i Haram'a ve ailesine selâmetle ulaştır." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Hacca giden hacımıza hayırlı yolculuklar ve kabul edilmiş bir hac nasip eyle. Onu Arafat, Müzdelife ve Mina'da ibadetle meşgul olan sâlih kullarından eyle. Beytullah'ı tavaf etmeyi, Safa ve Merve'de sa'y yapmayı nasip eyle. Haccını, duasını ve kurbanını makbul eyle. Onu günahlarından arınmış, temiz ve huzurlu olarak yurduna döndür. Okuduğumuz hatmin ve bu duanın sevabını onun ve bütün hacılarımızın üzerine hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const askerUgurlama = duaDoc("ASKER UĞURLAMA DUASI", "(Askere gideni uğurlarken okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "الْحَمْدُ لِلَّهِ رَبِّ الْعَالَمِينَ وَالصَّلَاةُ وَالسَّلَامُ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Elhamdülillâhi rabbi'l-âlemîn. Ves-salâtü ves-selâmü alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Hamd, âlemlerin Rabbi Allah'a mahsustur. Salât ve selâm, efendimiz Muhammed'in, âlinin ve ashâbının üzerine olsun." },
  { section: "Korunma Duası" },
  { ar: "بِسْمِ اللَّهِ الَّذِي لَا يَضُرُّ مَعَ اسْمِهِ شَيْءٌ فِي الْأَرْضِ وَلَا فِي السَّمَاءِ وَهُوَ السَّمِيعُ الْعَلِيمُ", okunus: "Bismillâhi'llezî lâ yedurru me'a'smihî şey'ün fi'l-erdı ve lâ fi's-semâi ve hüve's-semî'ul-alîm.", meal: "Adı anıldığında yerde ve gökte hiçbir şeyin zarar veremeyeceği Allah'ın adıyla. O, işitendir, bilendir." },
  { ar: "اللَّهُمَّ احْفَظْهُ مِنْ بَيْنِ يَدَيْهِ وَمِنْ خَلْفِهِ وَعَنْ يَمِينِهِ وَعَنْ شِمَالِهِ وَمِنْ فَوْقِهِ", okunus: "Allâhümmehfezhû min beyni yedeyhi ve min halfihî ve an yemînihî ve an şimâlihî ve min fevkih.", meal: "Allah'ım! Onu önünden, arkasından, sağından, solundan ve üstünden gelecek her türlü kötülükten koru." },
  { ar: "وَاللَّهُ خَيْرُ حَافِظًا وَهُوَ أَرْحَمُ الرَّاحِمِينَ", okunus: "Vallâhü hayru hâfizan ve hüve erhamü'r-râhimîn.", meal: "Allah koruyanların en hayırlısıdır ve O, merhamet edenlerin en merhametlisidir." },
  { section: "Vatan İçin Dua" },
  { ar: "اللَّهُمَّ احْفَظْ أَوْطَانَنَا وَاجْعَلْهَا أَمْنَاً آمِناً وَادْفَعْ عَنْهَا كُلَّ مَكْرُوهٍ", okunus: "Allâhümmehfaz evtânenâ ve'c'alhâ emnen âminen vedfa' anhâ külle mekrûh.", meal: "Allah'ım! Vatanımızı koru, onu güven içinde eyle ve ondan her türlü kötülüğü uzaklaştır." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Vatanı, milleti, bayrağı ve ezânı için askere giden bu kardeşimizi muhafaza eyle. Ona sıhhat, cesaret, dirayet ve sabır ihsan eyle. Silâh arkadaşlarıyla birlikte onu kazadan ve belâdan koru. Ailesini sabır ve huzurla rızıklandır. Onu sâlimen, âfiyetle ve şerefle vatanına, yuvasına döndür. Okuduğumuz hatmin ve bu duanın sevabını bütün askerlerimize, onların âilelerine ve vatanımızın bekâsına hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const isimKoyma = duaDoc("BEBEK İSİM KOYMA DUASI", "(Yedinci günde isim verilirken okunur; hatim programı kapsamında da okunur)", [
  { section: "İsim Koyma Âdâbı" },
  { meal: "İsim, yedinci günde verilir; sağ kulağa ezan, sol kulağa kâmet okunur, güzel ve anlamlı bir isim seçilir. Peygamberimiz (s.a.v.) güzel isim koymayı ve ismin güzel anlamını önemserdi." },
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "اللَّهُمَّ صَلِّ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Allâhümme salli alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Allah'ım! Efendimiz Muhammed'e, âline ve ashâbının tümüne rahmet eyle." },
  { section: "Âyet (İbrâhîm 37)" },
  { ar: "رَبِّ اجْعَلْ هَذَا الْبَلَدَ آمِنًا وَاجْنُبْنِي وَبَنِيَّ أَنْ نَعْبُدَ الْأَصْنَامَ", okunus: "Rabbi'c'al hâzâ'l-belede âminen vecnübniye ve beniyye en na'büde'l-esnâm.", meal: "Rabbim! Bu beldeyi güvenli kıl; beni ve oğullarımı putlara tapmaktan uzak tut." },
  { section: "İsim Duası" },
  { ar: "اللَّهُمَّ اجْعَلْهُ مُبَارَكًا وَاجْعَلْهُ صَالِحًا وَاجْعَلْهُ حَسَنَ الْخُلُقِ", okunus: "Allâhümme'c'alhû mübareken ve'c'alhû sâlihan ve'c'alhû hasene'l-huluk.", meal: "Allah'ım! Onu bereketli, sâlih ve güzel ahlâklı eyle." },
  { ar: "اللَّهُمَّ بَارِكْ فِيهِ وَفِي اسْمِهِ", okunus: "Allâhümme bârik fîhi ve fî'smih.", meal: "Allah'ım! Onu ve onun ismini bereketlendir." },
  { ar: "اللَّهُمَّ اجْعَلْهُ قُرَّةَ عَيْنٍ لِوَالِدَيْهِ وَبَرًّا بِهِمَا وَطَائِعًا لَكَ", okunus: "Allâhümme'c'alhû kurrete aynin li-vâlideyhi ve berren bihimâ ve tâian leke.", meal: "Allah'ım! Onu anne-babasına göz aydınlığı, onlara hayırlı, sana itaatkâr bir kul eyle." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Bu masum yavruya verilen ismi hayırlı ve mübarek eyle. Onu bu ismin anlamına yakışır, sâlih, âlim, takvâlı ve hayırlı bir kul eyle. Ona güzel ahlâk, sağlık ve uzun ömür ihsan eyle. Anne ve babasına, âilesine ve ümmetine faydalı eyle. Okuduğumuz hatmin ve bu duanın sevabını bu yavruya ve bütün masum çocuklara hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const seyyidulIstigfar = duaDoc("SEYYİDÜ'L-İSTİĞFAR", "(İstiğfarların efendisi — Sabah ve akşam okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "اللَّهُمَّ صَلِّ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Allâhümme salli alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Allah'ım! Efendimiz Muhammed'e, âline ve ashâbının tümüne rahmet eyle." },
  { section: "Seyyidü'l-İstiğfar" },
  { ar: "اللَّهُمَّ أَنْتَ رَبِّي لَا إِلَهَ إِلَّا أَنْتَ خَلَقْتَنِي وَأَنَا عَبْدُكَ وَأَنَا عَلَى عَهْدِكَ وَوَعْدِكَ مَا اسْتَطَعْتُ أَعُوذُ بِكَ مِنْ شَرِّ مَا صَنَعْتُ أَبُوءُ لَكَ بِنِعْمَتِكَ عَلَيَّ وَأَبُوءُ بِذَنْبِي فَاغْفِرْ لِي فَإِنَّهُ لَا يَغْفِرُ الذُّنُوبَ إِلَّا أَنْتَ", okunus: "Allâhümme ente Rabbî lâ ilâhe illâ ente halaktenî ve ene abdüke ve ene alâ ahdike ve va'dike mesteta'tü. Eûzü bike min şerri mâ sana'tü. Ebûü leke bi-ni'metike aleyye ve ebûü bi-zenbî. Fağfir lî feinnehû lâ yağfiru'z-zünûbe illâ ente.", meal: "Allah'ım! Sen benim Rabbimsin. Senden başka ilâh yoktur. Beni sen yarattın; ben senin kulunum. Gücüm yettiğince sana verdiğim söz ve ahdim üzereyim. İşlediğim şeylerin şerrinden sana sığınırım. Bana verdiğin nimeti itiraf eder, günahımı da itiraf ederim. Beni bağışla; çünkü günahları senden başka bağışlayacak yoktur." },
  { section: "Fazileti" },
  { meal: "Resûlullah (s.a.v.) buyurdu ki: \"Kim bu duayı gün boyunca ihlasla söyler de o gün akşam olmadan ölürse cennet ehlinden olur. Kim de gece ihlasla söyler de o gece sabah olmadan ölürse cennet ehlinden olur.\" (Buhârî, Deavât 2)" },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Bizleri, anne-babalarımızı, evlâtlarımızı, akraba ve bütün mü'minleri Seyyidü'l-istiğfar ile istiğfar etmeye muvaffak eyle. Günahlarımızı bağışla, kalplerimizi nurlandır, bedenlerimizi âfiyetlendir. Bu hatmi ve dualarımızı kabul buyur; bizi de ihlaslı istiğfar eden kullarından eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const kurbanDualari = duaDoc("KURBAN KESERKEN OKUNAN ÂYET VE DUALAR", "(Kurban keserken ve kesildiği anda okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "اللَّهُمَّ صَلِّ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Allâhümme salli alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Allah'ım! Efendimiz Muhammed'e, âline ve ashâbının tümüne rahmet eyle." },
  { section: "Kesime Başlarken" },
  { ar: "بِسْمِ اللَّهِ وَاللَّهُ أَكْبَرُ", okunus: "Bismillâh, vallâhü ekber.", meal: "Allah'ın adıyla, Allah en büyüktür." },
  { section: "Âyet (En'âm 162-163)" },
  { ar: "قُلْ إِنَّ صَلَاتِي وَنُسُكِي وَمَحْيَايَ وَمَمَاتِي لِلَّهِ رَبِّ الْعَالَمِينَ لَا شَرِيكَ لَهُ وَبِذَلِكَ أُمِرْتُ وَأَنَا أَوَّلُ الْمُسْلِمِينَ", okunus: "Kul inne salâtî ve nüsükî ve mahyâye ve memâtî lillâhi rabbi'l-âlemîn. Lâ şerîke lehû ve bizâlike ümirtü ve ene evvelü'l-müslimîn.", meal: "De ki: Şüphesiz benim namazım, ibadetim (kurbanım), hayatım ve ölümüm, âlemlerin Rabbi Allah içindir. O'nun ortağı yoktur. Bana bunu yapmam emrolundu ve ben müslümanların ilkiyim." },
  { section: "Âyet (En'âm 79)" },
  { ar: "إِنِّي وَجَّهْتُ وَجْهِيَ لِلَّذِي فَطَرَ السَّمَاوَاتِ وَالْأَرْضَ حَنِيفًا وَمَا أَنَا مِنَ الْمُشْرِكِينَ", okunus: "İnnî veccehtü vechiye lillezî fatara's-semâvâti ve'l-arda hanîfen ve mâ ene mine'l-müşrikîn.", meal: "Ben, hanîf (hakka yönelmiş) olarak yüzümü gökleri ve yeri yaratana çevirdim; ben müşriklerden değilim." },
  { section: "Âyet (Hac 37)" },
  { ar: "لَنْ يَنَالَ اللَّهَ لُحُومُهَا وَلَا دِمَاؤُهَا وَلَكِنْ يَنَالُهُ التَّقْوَى مِنْكُمْ", okunus: "Len yenâle'llâhe lühûmühâ ve lâ dimâühâ ve lâkin yenâlühü't-takvâ minküm.", meal: "Onların ne etleri ne de kanları Allah'a ulaşır; fakat O'na sizin takvânız ulaşır." },
  { section: "Kesim Duası" },
  { ar: "اللَّهُمَّ مِنْكَ وَلَكَ", okunus: "Allâhümme minke ve leke.", meal: "Allah'ım! Bu kurban sendendir, sanadır." },
  { ar: "اللَّهُمَّ تَقَبَّلْ مِنِّي كَمَا تَقَبَّلْتَ مِنْ إِبْرَاهِيمَ خَلِيلِكَ", okunus: "Allâhümme tekabbel minnî kemâ tekabbelte min İbrâhîme halîlik.", meal: "Allah'ım! Bu kurbanı, dostun İbrâhîm'den kabul ettiğin gibi benden de kabul eyle." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe)" },
  { meal: "Yâ Rab! Kestiğimiz kurbanı, kestiğimiz bu hayvanın etini, kanını ve fedakârlığımızı dergâhında kabul eyle. Bu kurbanın sevabını başta Peygamberimiz (s.a.v.), âilesi, ashâbı ve tüm inananların rûhlarına hediye eyledik, ulaştır. Bu ibadeti, takvâ sahiplerinin kurbanı gibi makbul eyle. Okuduğumuz hatmin sevabını, kesilen kurbanların sevabıyla birlikte bütün mü'minlerin rûhlarına hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

const nazarDuasi = duaDoc("NAZAR ÂYETİ VE DUASI", "(Nazar değmesine karşı okunur; hatim programı kapsamında da okunur)", [
  { section: "Hamd ü Senâ ve Salâvât" },
  { ar: "بِسْمِ اللَّهِ الرَّحْمَنِ الرَّحِيمِ", okunus: "Bismillâhi'r-rahmâni'r-rahîm.", meal: "Rahman ve Rahîm olan Allah'ın adıyla." },
  { ar: "اللَّهُمَّ صَلِّ عَلَى سَيِّدِنَا مُحَمَّدٍ وَعَلَى آلِهِ وَصَحْبِهِ أَجْمَعِينَ", okunus: "Allâhümme salli alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn.", meal: "Allah'ım! Efendimiz Muhammed'e, âline ve ashâbının tümüne rahmet eyle." },
  { section: "Nazar Âyeti (Kalem 51-52)" },
  { ar: "وَإِنْ يَكَادُ الَّذِينَ كَفَرُوا لَيُزْلِقُونَكَ بِأَبْصَارِهِمْ لَمَّا سَمِعُوا الذِّكْرَ وَيَقُولُونَ إِنَّهُ لَمَجْنُونٌ وَمَا هُوَ إِلَّا ذِكْرٌ لِلْعَالَمِينَ", okunus: "Ve in yekâdüllezîne keferû le-yüzlikûneke bi-ebsârihim lemmâ semiû'z-zikre ve yekûlûne innehû le-mecnûn. Ve mâ hüve illâ zikrun li'l-âlemîn.", meal: "Şüphesiz o inkâr edenler, Kur'ân'ı işittiklerinde neredeyse seni gözleriyle deviriverecekler; (senin için) 'O, mutlaka bir delidir' diyorlar. Oysa o (Kur'ân), âlemler için bir öğüttür." },
  { section: "Nazar Duası (Hadis-i Şerif'ten)" },
  { ar: "أَعُوذُ بِكَلِمَاتِ اللَّهِ التَّامَّاتِ مِنْ شَرِّ مَا خَلَقَ", okunus: "Eûzü bi-kelimâtillâhi't-tâmmeti min şerri mâ halek.", meal: "Allah'ın eksiksiz kelimelerine sığınırım; yarattıklarının şerrinden." },
  { ar: "أَعُوذُ بِكَلِمَاتِ اللَّهِ التَّامَّاتِ مِنْ كُلِّ شَيْطَانٍ وَهَامَّةٍ وَمِنْ كُلِّ عَيْنٍ لَامَّةٍ", okunus: "Eûzü bi-kelimâtillâhi't-tâmmeti min külli şeytânin ve hâmmetin ve min külli aynin lâmmetin.", meal: "Allah'ın eksiksiz kelimelerine sığınırım; her şeytandan, zehirli haşerattan ve isabet eden her gözden (nazardan)." },
  { section: "Güzel Söz" },
  { ar: "مَا شَاءَ اللَّهُ لَا قُوَّةَ إِلَّا بِاللَّهِ", okunus: "Mâ şâallâhü lâ kuvvete illâ billâh.", meal: "Allah'ın dilediği olur. Güç ve kuvvet ancak Allah'ındır." },
  { section: "Hatim Programı Kapsamında Kapanış (Türkçe dua)" },
  { meal: "Yâ Rab! Bizi ve sevdiklerimizi göz değmesinden, nazarın ve hasetçilerin şerrinden muhafaza eyle. Kur'ân-ı Kerîm'in ve Peygamberimizin (s.a.v.) duasının hürmetine, bizleri her türlü kötülükten, hastalıktan ve belâdan koru. Gönüllerimize iman, bedenlerimize sıhhat ve âilelerimize huzur ver. Okuduğumuz hatmin ve bu duanın sevabını bütün kardeşlerimize, çocuklarımıza ve âilelerimize hediye eyledik; kabul eyle. Âmîn, bi-hürmet-i seyyidi'l-mürselîn, ve'l-hamdü lillâhi rabbi'l-âlemîn." },
]);

function sureDoc(titleText, arabic, meal) {
  return [
    title(titleText),
    blank(),
    ...(Array.isArray(arabic) ? arabic : [arabic]).map((a) => p(a, { rtl: true, after: 200 })),
    blank(),
    section("MEÂLİ"),
    ...(Array.isArray(meal) ? meal : [meal]).map((m) => p(m, { italics: true, color: "44564e" })),
  ];
}

const ayetulKursi = sureDoc(
  "ÂYETÜ'L-KÜRSÎ (Bakara 255)",
  "اَللّٰهُ لَٓا اِلٰهَ اِلَّا هُوَ الْحَيُّ الْقَيُّومُۜ لَا تَأْخُذُهُ سِنَةٌ وَلَا نَوْمٌۜ لَهُ مَا فِي السَّمٰوَاتِ وَمَا فِي الْاَرْضِۜ مَنْ ذَا الَّذِي يَشْفَعُ عِنْدَهُٓ اِلَّا بِاِذْنِهٖۜ يَعْلَمُ مَا بَيْنَ اَيْد۪يهِمْ وَمَا خَلْفَهُمْۚ وَلَا يُح۪يطُونَ بِشَيْءٍ مِنْ عِلْمِهٖٓ اِلَّا بِمَا شَٓاءَۚ وَسِعَ كُرْسِيُّهُ السَّمٰوَاتِ وَالْاَرْضَۚ وَلَا يَؤُ۫دُهُ حِفْظُهُمَا وَهُوَ الْعَلِيُّ الْعَظ۪يمُ ۝٢٥٥",
  "Allah, kendisinden başka hiçbir ilâh olmayandır. Diridir, her şeyi ayakta tutandır. O'nu ne bir uyuklama tutabilir ne de bir uyku. Göklerdeki her şey, yerdeki her şey O'nundur. İzni olmaksızın O'nun katında şefaat edecek kimdir? O, kulların önlerindekini ve arkalarındakini bilir. Onlar, O'nun ilminden dilediği kadarından başkasını kavrayamazlar. O'nun kürsüsü gökleri ve yeri kaplamıştır. Onları koruyup gözetmek O'na güç gelmez. O yücedir, büyüktür."
);

const hasr = sureDoc(
  "HAŞR SÛRESİ SON ÜÇ ÂYET (59/22-24)",
  [
    "اَللّٰهُ الَّذ۪ي لَٓا اِلٰهَ اِلَّا هُوَۚ عَالِمُ الْغَيْبِ وَالشَّهَادَةِۚ هُوَ الرَّحْمٰنُ الرَّح۪يمُ ۝٢٢",
    "هُوَ اللّٰهُ الَّذ۪ي لَٓا اِلٰهَ اِلَّا هُوَۚ اَلْمَلِكُ الْقُدُّوسُ السَّلَامُ الْمُؤْمِنُ الْمُهَيْمِنُ الْعَز۪يزُ الْجَبَّارُ الْمُتَكَبِّرُۜ سُبْحَانَ اللّٰهِ عَمَّا يُشْرِكُونَ ۝٢٣",
    "هُوَ اللّٰهُ الْخَالِقُ الْبَارِئُ الْمُصَوِّرُ لَهُ الْاَسْمَٓاءُ الْحُسْنٰىۜ يُسَبِّحُ لَهُ مَا فِي السَّمٰوَاتِ وَالْاَرْضِۚ وَهُوَ الْعَز۪يزُ الْحَك۪يمُ ۝٢٤",
  ],
  [
    "O, kendisinden başka ilâh olmayan Allah'tır. Gaybı da görünen âlemi de bilir. O, Rahmân'dır, Rahîm'dir.",
    "O, mülkün sahibi, kutsal, selâmet veren, güven veren, gözeten, üstün, istediğine gücü yeten, büyüklükte eşi olmayan Allah'tır. Allah, onların ortak koştuklarından münezzehtir.",
    "O, yaratan, var eden, şekil veren Allah'tır. En güzel isimler O'nundur. Göklerde ve yerde olanlar O'nu tesbih eder. O, güçlüdür, hikmet sahibidir.",
  ]
);

const ahzab = sureDoc(
  "AHZÂB SÛRESİ 56. ÂYET (Salâvat Âyeti)",
  "اِنَّ اللّٰهَ وَمَلٰٓئِكَتَهُ يُصَلُّونَ عَلَى النَّبِيِّۜ يَٓا اَيُّهَا الَّذ۪ينَ اٰمَنُوا صَلُّوا عَلَيْهِ وَسَلِّمُوا تَسْل۪يمًا ۝٥٦",
  "Şüphesiz Allah ve melekleri Peygamber'e salât ederler. Ey iman edenler! Siz de ona salât edin, selâm edin."
);

const fetih = sureDoc(
  "FETİH SÛRESİ 29. ÂYET",
  "مُحَمَّدٌ رَسُولُ اللّٰهِۜ وَالَّذ۪ينَ مَعَهُٓ اَشِدَّٓاءُ عَلَى الْكُفَّارِ رُحَمَٓاءُ بَيْنَهُمْ تَرٰيهُمْ رُكَّعًا سُجَّدًا يَبْتَغُونَ فَضْلًا مِنَ اللّٰهِ وَرِضْوَانًاۘ س۪يمَاهُمْ ف۪ي وُجُوهِهِمْ مِنْ اَثَرِ السُّجُودِۜ ۝٢٩",
  "Muhammed, Allah'ın Resûlü'dür. Onunla beraber olanlar, kâfirlere karşı şiddetli, kendi aralarında merhametlidirler. Onları rükû eden, secde eden, Allah'tan lütuf ve rızâ dileyen kimseler olarak görürsün. Secdelerin izi yüzlerindedir."
);

const zumer = sureDoc(
  "ZÜMER SÛRESİ 53. ÂYET",
  "قُلْ يَا عِبَادِيَ الَّذ۪ينَ اَسْرَفُوا عَلٰٓى اَنْفُسِهِمْ لَا تَقْنَطُوا مِنْ رَحْمَةِ اللّٰهِۜ اِنَّ اللّٰهَ يَغْفِرُ الذُّنُوبَ جَم۪يعًاۜ اِنَّهُ هُوَ الْغَفُورُ الرَّح۪يمُ ۝٥٣",
  "De ki: \"Ey kendilerine zulmederek aşırı giden kullarım! Allah'ın rahmetinden ümit kesmeyin. Şüphesiz Allah bütün günahları bağışlar. O, çok bağışlayandır, çok merhamet edendir.\""
);

const rad = sureDoc(
  "RA'D SÛRESİ 28. ÂYET",
  "اَلَّذ۪ينَ اٰمَنُوا وَتَطْمَئِنُّ قُلُوبُهُمْ بِذِكْرِ اللّٰهِۜ اَلَا بِذِكْرِ اللّٰهِ تَطْمَئِنُّ الْقُلُوبُ ۝٢٨",
  "Onlar, iman edenler ve gönülleri Allah'ı anmakla huzura kavuşanlardır. Biliniz ki kalpler ancak Allah'ı anmakla huzur bulur."
);

const mulk = sureDoc(
  "MÜLK (TEBÂREKE) SÛRESİ BAŞI (1-4)",
  [
    "تَبَارَكَ الَّذ۪ي بِيَدِهِ الْمُلْكُ وَهُوَ عَلٰى كُلِّ شَيْءٍ قَد۪يرٌۙ ۝١",
    "الَّذ۪ي خَلَقَ الْمَوْتَ وَالْحَيٰوةَ لِيَبْلُوَكُمْ اَيُّكُمْ اَحْسَنُ عَمَلًاۜ وَهُوَ الْعَز۪يزُ الْغَفُورُۙ ۝٢",
    "الَّذ۪ي خَلَقَ سَبْعَ سَمٰوَاتٍ طِبَاقًاۜ مَا تَرٰى ف۪ي خَلْقِ الرَّحْمٰنِ مِنْ تَفَاوُتٍۜ فَارْجِعِ الْبَصَرَۙ هَلْ تَرٰى مِنْ فُطُورٍ ۝٣",
    "ثُمَّ ارْجِعِ الْبَصَرَ كَرَّتَيْنِ يَنْقَلِبْ اِلَيْكَ الْبَصَرُ خَاسِئًا وَهُوَ حَس۪يرٌ ۝٤",
  ],
  [
    "Hükümranlık elinde olan Allah ne yücedir. O, her şeye hakkıyla gücü yetendir.",
    "O, hanginizin daha güzel amel yapacağını sınamak için ölümü ve hayatı yaratandır. O, mutlak güç sahibidir, çok bağışlayandır.",
    "O, yedi göğü tabaka tabaka yaratandır. Rahmân'ın yaratışında hiçbir uyumsuzluk göremezsin. Gözünü çevir de bak, bir kusur görebilir misin?",
    "Sonra gözünü tekrar tekrar çevir bak. Göz, aradığı bir kusur bulamayıp usanmış ve bitkin olarak sana döner.",
  ]
);

function mevlidDoc(titleText, stanzas) {
  const body = [title(titleText), blank()];
  stanzas.forEach((s) => {
    body.push(section(s[0]));
    s.slice(1).forEach((line) => body.push(p(line, { align: AlignmentType.CENTER })));
    body.push(blank());
  });
  return body;
}

const tevhid = mevlidDoc("MEVLİD-İ ŞERÎF — TEVHİD (BAŞLANGIÇ) BAHRİ", [
  [
    "Tevhid Bahri",
    "Allah adın zikredelim evvelâ",
    "Vâcib oldur cümle işte her kula",
    "Allah adın her kim ol evvel ana",
    "Her işi âsân eder Allah ana",
    "Bir kez Allah dese aşk ile lisân",
    "Dökülür cümle günâh misl-i hazân",
    "Hakk'a âsî kullara rahmet gelür",
    "Bu murâd üzre murâda erdürür",
  ],
]);

const veladet = mevlidDoc("MEVLİD-İ ŞERÎF — VELÂDET (NÛR) BAHRİ", [
  [
    "Velâdet (Nûr) Bahri",
    "Âmine hâtun Muhammed anası",
    "Ol sadefden doğdı ol dürr-i hüdâ",
    "Çün dokuz ay tamâm oldı gece",
    "Seher ânında bu resûl geldi ne",
    "Evvel âhir ol didi Allah'un adı",
    "Hem Muhammed Peygamber'un anun adı",
    "Düşdi nûr-ı Mustafâ dört yanına",
    "Gitdi küfür zulmeti îmân yanına",
    "Yer gök nûr ile toldı arş-ı Hudâ",
    "Nûr-ı Muhammed'den oldı pür ziyâ",
    "Kim ki bir kez gördü anun cemâlin",
    "İster idi görmeğe dâim cemâlin",
  ],
]);

const mirac = mevlidDoc("MEVLİD-İ ŞERÎF — MİRAÇ BAHRİ", [
  [
    "Miraç Bahri (seçme)",
    "Hak anı Mi'râca okudı gece",
    "Geldi Cibrîl-i emîn didi hoca",
    "Geldi Cebrâil didi yâ Mustafâ",
    "Kıl urûc ehl-i semâ gör sen bizi",
    "Arş-ı a'lâda işit sen sırr-ı râzı",
    "Gör resûlün hâlini ey şâh-ı gâzî",
  ],
]);

const rihlet = mevlidDoc("MEVLİD-İ ŞERÎF — RİHLET (VEFAT) BAHRİ", [
  [
    "Rihlet (Vefat) Bahri — giriş",
    "Şol rivâyet kim anun vefâtınun",
    "Ben hem anun degdürin cevâbınun",
    "Aceb oldı işitdüm bu haberi",
    "Beni dutdı bu sözün yakıcı nârı",
  ],
]);

const salevat = [
  title("SALÂVÂT-I ŞERÎFE"),
  blank(),
  p("Mevlid ve cemiyet programlarında en çok okunan salâvâtlar şunlardır. Salâvât getirirken cemaat da 'Allâhümme salli alâ seyyidinâ Muhammed' diyerek iştirak eder."),
  blank(),
  section("1. Kısa Salâvât"),
  p("Allâhümme salli alâ seyyidinâ Muhammedin ve alâ âli seyyidinâ Muhammed."),
  section("2. Salavât-ı Münciye (Cevşen salâvâtı)"),
  p("Allâhümme salli alâ seyyidinâ Muhammedin ve alâ âlihî adede mâ fî ilmillâhi salâten dâimeten bi-devâmi mülkillâh."),
  section("3. Salavât-ı Fetih (Tefriciye - Dua-yı Nâriye)"),
  p("Allâhümme salli salâten kâmileten ve sellim selâmen tâmmen alâ seyyidinâ Muhammedin ellezî tenhallü bihil-ukadü ve tenfericü bihil-kürebü ve tukzâ bihil-havâicu ve tünâlü bihir-reğâibü ve hüsnül-havâtimi ve yüsteskal-ğamâmu bi-vechihil-kerîm ve alâ âlihî ve sahbihî fî külli lemhatin ve nefesin bi-adedi külli ma'lûmin lek."),
  section("4. Salavât-ı Şifâ"),
  p("Allâhümme salli alâ seyyidinâ Muhammedin tabîbi'l-kulûbi ve devâihâ ve âfiyeti'l-ebdâni ve şifâihâ ve nûri'l-ebsâri ve dıyâihâ ve alâ âlihî ve sahbihî ve sellim."),
  section("5. Mevlid sonunda okunan salâvât"),
  p("Allâhümme salli alâ seyyidinâ Muhammedin ve alâ âlihî ve sahbihî ecmaîn."),
  blank(),
  p("Not: Salâvâtların fazileti hakkında güvenilir kaynaklardan bilgi alınabilir. Metinler yaygın kullanılan şekilleriyle yazılmıştır; bazı varyantlar olabilir.", { italics: true, color: "44564e" }),
];

const ilahiler = [
  title("MEVLİD İLÂHİLERİ"),
  blank(),
  p("Mevlid bahrları arasında ve program sonunda okunan meşhur ilâhilerden örnekler. Bazı ilâhilerin metinleri yörelere göre farklılık gösterebilir."),
  blank(),
  section("1. Uyan ey gözlerim gafletten uyan"),
  p("Uyan ey gözlerim gafletten uyan", { align: AlignmentType.CENTER }),
  p("Uyan uykusu çok gözlerim uyan", { align: AlignmentType.CENTER }),
  p("Kıl namazın, kurtulur sanın", { align: AlignmentType.CENTER }),
  p("Uyan uykusu çok gözlerim uyan", { align: AlignmentType.CENTER }),
  blank(),
  section("2. Safâ Dağı"),
  p("Safâ dağı sırtında", { align: AlignmentType.CENTER }),
  p("Bir çoban gezer idi", { align: AlignmentType.CENTER }),
  p("Resûlün aşkı ile", { align: AlignmentType.CENTER }),
  p("Yüreği yanar idi", { align: AlignmentType.CENTER }),
  blank(),
  section("3. Allah Allah ilâhi"),
  p("Allah Allah ilâhi", { align: AlignmentType.CENTER }),
  p("Senin adın yâ Hû", { align: AlignmentType.CENTER }),
  p("Senin dergâhına kulluk", { align: AlignmentType.CENTER }),
  p("Meylimiz yâ Hû", { align: AlignmentType.CENTER }),
  blank(),
  section("4. Ya Muhammed (Salâvat ilâhisi)"),
  p("Ya Muhammed canım arzular seni", { align: AlignmentType.CENTER }),
  p("İki cihânın mehri olan Resûl", { align: AlignmentType.CENTER }),
  blank(),
  section("5. Bir Gece Bir Gece"),
  p("Bir gece bir gece", { align: AlignmentType.CENTER }),
  p("Melekler bir gece", { align: AlignmentType.CENTER }),
  p("Semadan inerler", { align: AlignmentType.CENTER }),
  p("Yeryüzüne bir gece", { align: AlignmentType.CENTER }),
  blank(),
  p("Not: İlâhi metinleri geleneksel sözlü kültürden derlenmiştir. Camii ve cemiyetinizde kullanılan asıl metinlerle karşılaştırıp kullanmanız tavsiye edilir.", { italics: true, color: "44564e" }),
];

const kasideler = [
  title("KASİDELER — GENEL BİLGİ"),
  blank(),
  p("Kaside, Arap ve Türk edebiyatında bir büyüğü, peygamberi veya dini bir şahsiyeti övmek, ona yalvarmak için yazılan manzum eserlerdir. Dini programlarda özellikle Mevlid bahrlarından sonra na't, münâcât ve salâvat üslûbunda kasideler okunur."),
  blank(),
  section("Meşhur Kasideler ve Na'tlar"),
  p("• Kasîde-i Bürde — İmam Bûsîrî (Hz. Peygamber'e yazılmış en meşhur kaside)"),
  p("• Kasîde-i Bür'ê — Ka'b bin Züheyr (Bânet Suâd)"),
  p("• Na't-ı Şerîf — Mevlâna Celâleddîn-i Rûmî ve diğer divan şairlerimizin na'tları"),
  p("• Münâcâtlar — Yûnus Emre, Süleyman Çelebi ve çeşitli şairlerin Allah'a yalvarış manzumeleri"),
  blank(),
  section("Programlarda kullanımı"),
  p("Kasideler genellikle makamla okunur; koro ve dinleyiciler 'salâvat' ile eşlik eder. Uzun kasidelerin tamamı yerine bahr bahr seçmeler okunabilir."),
  p("Kendi cemaatinizin geleneğindeki kasideleri bu klasöre Word veya PDF olarak ekleyebilirsiniz.", { italics: true, color: "44564e" }),
];

const burde = [
  title("KASÎDE-İ BÜRDE — GİRİŞ BEYİTLERİ"),
  section("(İmam Bûsîrî)"),
  blank(),
  section("Arapça metin"),
  p("أَمِنْ تَذَكُّرِ جِيرَانٍ بِذِي سَلَمِ", { rtl: true }),
  p("مَزَجْتَ دَمْعًا جَرَى مِنْ مُقْلَةٍ بِدَمِ", { rtl: true }),
  p("أَمْ هَبَّتِ الرِّيحُ مِنْ تِلْقَاءِ كَاظِمَةٍ", { rtl: true }),
  p("وَأَوْمَضَ الْبَرْقُ فِي الظَّلْمَاءِ مِنْ إِضَمِ", { rtl: true }),
  p("فَمَا لِعَيْنَيْكَ إِنْ قُلْتَ اكْفُفَا هَمَتَا", { rtl: true }),
  p("وَمَا لِقَلْبِكَ إِنْ قُلْتَ اسْتَفِقْ يَهِمِ", { rtl: true }),
  blank(),
  section("Anlam özeti"),
  p("'Selam köyündeki sevgilileri hatırlamaktan mı gözünden kan karışık yaşlar akıttın? Yoksa Kâzıme taraflarından rüzgâr mı esti, İdam'dan gece karanlığında şimşek mi çaktı? Gözlerine ne oldu; 'dur' desen de ağlıyor, kalbine ne oldu; 'ayıl' desen de dalıp gidiyor?'", { italics: true, color: "44564e" }),
  blank(),
  p("Not: Kasîde-i Bürde tam metni uzundur. Bu dosya yalnızca giriş beyitlerini içerir; tam metin ve meali için güvenilir bir mevlit/kaside kitabından yararlanabilirsiniz.", { italics: true, color: "44564e" }),
];

function ottDoc(titleText, subtitle, couplets) {
  const body = [
    title(titleText),
    p(subtitle || "Osmanlica - Arap harfli metin", { italics: true, align: AlignmentType.CENTER, color: "44564e", size: 22 }),
    blank(),
  ];
  couplets.forEach((pair) => {
    body.push(p(pair[0], { rtl: true, align: AlignmentType.CENTER, after: 40 }));
    body.push(p(pair[1], { rtl: true, align: AlignmentType.CENTER, after: 160 }));
  });
  return body;
}

const ottTanitim = [
  title("MEVLİD-İ ŞERÎF — OSMANLICA (HAREKELİ) NÜSHASI"),
  blank(),
  section("Bu klasördeki dosyalar"),
  p("Süleyman Çelebi'nin Vesîletü'n-Necât adlı mevlidinin Arap harfli (Osmanlıca), harekeli (harekeli) metinleri bu klasörde bahr bahr ayrı ayrı Word dosyaları olarak durur:"),
  p("• Tevhid-Bahri — Allah adın zikredelim evvelâ"),
  p("• Veladet-Bahri — Peygamberimizin (s.a.v.) doğumu"),
  p("• Amine-Hatun-Bahri — Âmine hâtun Muhammed ânesi"),
  p("• Merhaba-Bahri — Merhabâ ey âlî sultân"),
  p("• Mirac-Bahri — Mi'rac yolculuğu"),
  p("• Rihlet-Bahri — Vefat-ı Nebevî"),
  p("• Dua-Iltica-Bahri — Yâ ilâhî sakla-gıl îmânımız"),
  blank(),
  section("Hareke (üstün-esre-ötre) hakkında"),
  p("Osmanlıca metinler, okunuşu kolaylaştırmak için üstün (َ), esre (ِ), ötre (ُ), cezim (ْ) ve şedde (ّ) işaretleriyle birlikte yazılmıştır. Türkçe sesli harfler için şu işaretler kullanılır: a/â = üstün, e = üstün, ı/i = esre, u/ü = ötre, o/ö = vav ile ötre."),
  blank(),
  section("Önemli not"),
  p("Bu metinler, ilmî neşirlerdeki Latin harfli metin esas alınarak ortak Osmanlı imlasıyla yazıya geçirilmiş ve harekelendirilmiştir. Mevlid'in el yazması nüshalarında küçük imla ve kelime farklılıkları bulunur. Programlarda okunacaksa güvenilir bir Osmanlıca mevlid baskısıyla karşılaştırmanız tavsiye edilir. Mevlid'in tamamı (700'e yakın beyit) bu klasördeki seçme beyitlerden uzundur; tam metin için bir Osmanlıca mevlid kitabından yararlanıp buraya ekleyebilirsiniz.", { italics: true, color: "44564e" }),
];

const ottTevhid = [
  ["اَللّٰه آدِن ذِکْر اِیدَلُم اَوَّلَا", "وَاجِب اُولْدُر جُملَه اِیشْدِه هَرْ قُلَه"],
  ["اَللّٰه آدِن هَرْ کِم اُول اَوَّل آنا", "هَرْ اِیشِی آسَان اِیدِه اَللّٰه آنا"],
  ["اَللّٰه آدِی اُولْسَه هَرْ اِیشِڭ اُوکِی", "هَرْگِز اَبْتَر اُولْمَیَه اَنِڭ سُوکِی"],
  ["هَرْ نَفَسْدِه اَللّٰه آدِن دِی مُدَام", "اَللّٰه آدِییْلَه اُولُر هَرْ اِیش تَمَام"],
  ["بِرْ کِز اَللّٰه دِیْسِه عَشْق اِیلِه لِسَان", "دُکِیلُر جُملَه گُنَاه مِثْلِ خَزَان"],
  ["حَقَّه عَاصِی قُلَّرَه رَحْمَت کِلُر", "بُو مُرَاد اُوزْرِه مُرَادَه اِرْدِیُر"],
  ["خَلْقِ عَالَم چُون کُرُرْلَر آدِنِی", "ذِکْر اِیدَرْلَر دَایِمَا اَللّٰه آدِنِی"],
];

const ottVeladet = [
  ["آمِنِه خَاتُون مُحَمَّد آنَاسِی", "اُول صَدَفْدَن طُوغْدِی اُول دُرِّ هُدَا"],
  ["چُون طُقُوز آی تَمَام اُولْدِی کِجِه", "سَحَر آنِنْدِه بُو رَسُول کِلْدِی نِه"],
  ["اُول کِجِه کِه طُوغْدِی اُول خَیْرُ الْبَشَر", "آنَاسِی مَاهِنْدِه کُورْدِی چُوق هُنَر"],
  ["اَوَّل آخِر اُول دِیدِی اَللّٰهِڭ آدِی", "هَم مُحَمَّد پَیْغَمْبَرِڭ اَنِڭ آدِی"],
  ["دُوشْدِی نُورِ مُصْطَفٰی دُرْت یَانِیْنَہ", "کِیتْدِی کُفْر ظُلْمَتِی اِیمَان یَانِیْنَہ"],
  ["یِرْ گُوکْ نُور اِیلِه طُولْدِی عَرْشِ خُدَا", "نُورِ مُحَمَّدْدَن اُولْدِی پُرْ ضِیَا"],
  ["چُون طُوغُوبْدُر اُول مُحَمَّد مُصْطَفٰی", "گِرْچِی کِه طُوغْدِی کِلْدِی آیِت بُو جَا"],
];

const ottMerhaba = [
  ["یَارَدِیلْمِش جُملَه اُولْدِی شَادُمَان", "غَم کِدِیُوب عَالَم یِڭیدن بُلْدِی جَان"],
  ["جُملَه ذَرَّاتِ جِهَان اِیدِیُوب نِدَا", "چَاغْرِشُوبَن دِیدِیلَر کِه مَرْحَبَا"],
  ["مَرْحَبَا اِی عَالِی سُلْطَان مَرْحَبَا", "مَرْحَبَا اِی کَانِ عِرْفَان مَرْحَبَا"],
  ["مَرْحَبَا اِی سِرِّ فُرْقَان مَرْحَبَا", "مَرْحَبَا اِی نُورِ رَحْمَن مَرْحَبَا"],
  ["مَرْحَبَا اِی بُلْبُلِ بَاغِ جَمَال", "مَرْحَبَا اِی آشِنَايِ ذِی الْجَلَال"],
  ["مَرْحَبَا اِی جَانِ بَاقِی مَرْحَبَا", "مَرْحَبَا اِی عُشَّاقَه سَاقِی مَرْحَبَا"],
  ["مَرْحَبَا اِی جَانِ جَانَان مَرْحَبَا", "مَرْحَبَا اِی دَرْدِه دَرْمَان مَرْحَبَا"],
  ["مَرْحَبَا اِی جُملِنِڭ مَطْلُوبِی سِنْ", "مَرْحَبَا اِی خَالِقِنِڭ مَحْبُوبِی سِنْ"],
  ["مَرْحَبَا اِی پَادْشَاهِ دُو جِهَان", "سِنِڭ اِیچُون اُولْدِی کَوْن اِیلِه مَکَان"],
  ["مَرْحَبَا اِی رَحْمَةً لِلْعَالَمِین", "مَرْحَبَا سِنْسِڭ شَفِیعُ الْمُذْنِبِین"],
  ["اِی گُوكُلْلَر دَرْدِنِڭ دَرْمَانِی سِنْ", "اِی یَارَدِیلْمِشْلَرِنِڭ سُلْطَانِی سِنْ"],
  ["سِنْسِڭ اُول سُلْطَانِ جُملَه اَنْبِیَا", "نُورِ چَشْمِ اَوْلِیَا وُ اَصْفِیَا"],
  ["یَا حَبِیب اللّٰه بِیزِه اِمْدَاد قِیل", "سُوڭ نَفَس دِیدَارِڭ اِیلِه شَاد قِیل"],
];

const ottMirac = [
  ["حَق آنِی مِعْرَاجَه اُوخُودِی کِجِه", "کِلْدِی جِبْرِیل اَمِین دِیدِی خُوَاجَه"],
  ["قِیل عُرُوج اَهْلِ سَمَا کُور سِنْ بِیزِی", "عَرْشِ اَعْلَادَه اِیشِیت سِنْ سِرِّ رَازِی"],
];

const ottRihlet = [
  ["شُول رِوَایَت کِه اَنِڭ وِفَاتِنِڭ", "بِنْ هَم اَنِڭ دِگْدِیْرِن جَوَابِنِڭ"],
  ["عَجَب اُولْدِی اِیشِیتْدُم بُو خَبَرِی", "بِنِی طُوتْدِی بُو سُوزِڭ یَاقِیجِی نَارِی"],
];

const ottAmine = [
  ["آمِنِه خَاتُون مُحَمَّد آنَاسِی", "اُول صَدَفْدَن طُوغْدِی اُول دُرِّ دَانَاسِی"],
  ["چُون کِه عَبْدُ اللّٰه دَن اُولْدِی حَامِلَه", "وَقْت اِیرِشْدِی هَفْتَه وُ اَیَّام اِیلِه"],
  ["هَم مُحَمَّد گِلْمِسِی اُولْدِی یَقِین", "چُوق عَلاَمَت لَر بِلِیرْدِی گِلْمَدِن"],
  ["اُول کِجِه کِه طُوغْدِی اُول خَیْرُ الْبَشَر", "آنَاسِی اَنْدَه نِلَر کُورْدِی نِلَر"],
  ["دِیدِی کُورْدُم اُول حَبِیبِڭ آنَاسِی", "بِرْ عَجَب نُور کِه گُوكِشْ پَرْوَانَاسِی"],
  ["بِرْک اُورُوب چِیقْدِی بِیْتِمْدَن نَاجَهَان", "گُوکْلَرَه دَک نُور اِیلِه طُولْدِی جِهَان"],
  ["گُوکْلَر آچِلْدِی و فَتْح اُولْدِی ظُلَم", "اُوچ مِلَک کُورْدُم اِلِنْدِه اُوچ عَلَم"],
  ["بِرِی مَشْرِق بِرِی مَغْرِبْدَه اَنِڭ", "بِرِی دَامِنْدَه دِکِیلْدِی کَعْبِنِڭ"],
  ["اِنْدِیلَر گُوکْدَن مِلَکْلَر صَاف صَاف", "کَعْبَه کِبِی قِلْدِیلَر اِوِم طَوَاف"],
  ["چَوْرَه یَانِمَه گِلُوب اُتُورْدِیلَر", "مُصْطَفٰایِ بِرْبِرِنَه مُشْتُولَر"],
  ["دِیدِیلَر اُوغْلُڭ کِبِی هِیچ بِر اُوغُل", "یَارَدِلَالِی جِهَان گِلْمِش دِگِل"],
  ["بُو سِنِڭ اُوغْلُڭ کِبِی قَدْرِی جَمِیل", "بِرْ آنَایَه وَرْمِمِشْدِر اُول جَلِیل"],
  ["اُولُو دَوْلَت بُلْدُڭ اِی دِلْدَار سِنْ", "طُوغِیسِیْسَرْدِر سِنْدَن اُول خُلْقِی حَسَن"],
  ["آمِنِه اِیدِر چُو وَقْت اُولْدِی تَمَام", "کِه وُجُودَه کِلَه اُول خَیْرُ الْاَنَام"],
  ["سُوسَدُمْ گَایَت حَرَارَت دَن قَتِی", "سُنْدِیلَر بِرْ جَام طُولُوسِی شَرْبَتِی"],
  ["شَرْبَتِی سُنْدِی طَابَانَه حُورِیلَر", "بُوڭُو سَانَه وَرْدِی اَللّٰه دِیدِیلَر"],
  ["کَارْدَن آق اِیدِی و هَم صُوغُق اِیدِی", "لَذَّتِی دَاهِی شَکَرْدَه یُوخ اِیدِی"],
  ["اِیچْدِم آنِی اُولْدِی جِسْمِم نُورَه غَرْق", "اِیدَمِزْدِم کَنْدِمِی نُورْدَن فَرْق"],
  ["گِلْدِی بِرْ آقْقُوش قَانَادِیلَه رَوَان", "آرْقَمِی صِیغَادِی قُوْوَتْلَه هَمَان"],
  ["طُوغْدِی اُول سَاعَتْدَه اُول سُلْطَانِ دِین", "نُورَه غَرْق اُولْدِی سَمَاوَات و زَمِین"],
];

const ottDuaIltica = [
  ["اِی عَزِیزْلَر اُوشْدَه بَاشْلَرُز سُوزَه", "بِرْ وَصِیَّت قِلَارُز اِلَّا سِیزَه"],
  ["اُول وَصِیَّت کِه دِرَم هَرْ کِم طُوتَه", "مِشْک کِبِی قُوْقُوسِی جَانْلَرْدَه طُوتَه"],
  ["حَق تَعَالٰی رَحْمَت اِیلِیَه آنا", "کِه بِنی اُول بِرْ دُعَا اِیلِه آنا"],
  ["هَرْ کِی دِیلَر بُو دُعَادَه بُلْمَه", "فَاتِحَه اِحْسَان اِیدِه بِن قُلُومَه"],
  ["یَا اِلٰهِی صَقْلَاغِل اِیمَانِمِز", "وِیرَلِم اِیمَان اِیلِه طَا جَانِمِز"],
  ["بِیز گُنَه کَار عَاصِی مُجْرِم قُلَّرِی", "یَرْلِیغَیُوب قِل گُونَاه لَرْدَن بَرِی"],
  ["قَبْرِمِز اِیمَان اِیلِه پُرْ نُور قِل", "مُونِسِڭ غِلْمَان اِیلِه هَم حُور قِل"],
  ["هَم دَاهِی مِیزَانِمِز اِیلِه سَقِیل", "جَنَّتَه گِیرْمِیَه لُطْفِڭ قِل دَلِیل"],
  ["مُصْطَفٰایَه جِوَار اِیت یَا کَرِیم", "جَنَّتُ الْفِرْدَوْس اِیچِنْدَه یَا رَحِیم"],
  ["لُطْف اِیلِه کُوسْتِر بِیزَه دِیدَارِنِی", "نِعْمَتِڭ لَه طُوْیْلَاغِل قُلَّرِنِی"],
  ["عَفْو اِیدُوب اِسْیَانِمِز قِل رَحْمَتِی", "اُول حَبِیبِڭ یُوزِی سُویِی حُرْمَتِی"],
  ["سَانَه لَایِق قُلَّر اِیلِه هَمْ دَم اِیت", "اَهْلِ دَرْدِڭ صُوحْبَتِنَه مَحْرَم اِیت"],
  ["هَم سُلَیْمَان فَقِیرَه رَحْمَت اِیت", "یُولْدَاشِڭ اِیمَان مَقَامِڭ جَنَّت اِیت"],
  ["یَا اِلٰهِی قِلْمَه بِیزِی دَالِّین", "بُو دُعَایَه جُملِڭِز دِییِڭ آمِین"],
];

/* ================= URETIM ================= */

(async () => {
  console.log("Dini Kutuphane ornek icerigi uretiliyor...\n");
  await save("Dualar", "Program_Duasi", programDuasi);
  await save("Dualar", "Hatim_Duasi", hatimDuasi);
  await save("Dualar", "Vaaz-Baslama-Duasi", vaazBaslama);
  await save("Dualar", "Cenaze-Duasi-Erkek", cenazeErkek);
  await save("Dualar", "Cenaze-Duasi-Kadin", cenazeKadin);
  await save("Dualar", "Cenaze-Duasi-Cocuk", cenazeCocuk);
  await save("Dualar", "Kabir-Ziyareti-Duasi", kabirZiyareti);
  await save("Dualar", "Kabir-Ziyaretinde-Yapilacak-Dua", kabirDua);
  await save("Dualar", "Telkin-Duasi", telkin);
  await save("Dualar", "Nisan-Duasi", nisanDuasi);
  await save("Dualar", "Nikah-Duasi", nikahDuasi);
  await save("Dualar", "Gelin-Ugurlama-Duasi", gelinUgurlama);
  await save("Dualar", "Yeni-Dogan-Cocuk-Duasi", yeniDogan);
  await save("Dualar", "Sunnet-Duasi", sunnetDuasi);
  await save("Dualar", "Pazar-Duasi", pazarDuasi);
  await save("Dualar", "Isyeri-Acilis-Temel-Atma-Duasi", isyeriAcilis);
  await save("Dualar", "Yagmur-Duasi", yagmurDuasi);
  await save("Dualar", "Umre-Ugurlama-Duasi", umreUgurlama);
  await save("Dualar", "Haci-Ugurlama-Duasi", haciUgurlama);
  await save("Dualar", "Asker-Ugurlama-Duasi", askerUgurlama);
  await save("Dualar", "Bebek-Isim-Koyma-Duasi", isimKoyma);
  await save("Dualar", "Seyyidul-Istigfar", seyyidulIstigfar);
  await save("Dualar", "Kurban-Keserken-Okunan-Dualar", kurbanDualari);
  await save("Dualar", "Nazar-Duasi", nazarDuasi);
  await save("Sureler", "Ayetul-Kursi", ayetulKursi);
  await save("Sureler", "Hasr-Suresi-Son-Uc-Ayet", hasr);
  await save("Sureler", "Ahzab-56-Salavat-Ayeti", ahzab);
  await save("Sureler", "Fetih-29", fetih);
  await save("Sureler", "Zumer-53", zumer);
  await save("Sureler", "Rad-28", rad);
  await save("Sureler", "Mulk-Suresi-Basi", mulk);
  await save("Mevlid", "Tevhid-Bahri", tevhid);
  await save("Mevlid", "Veladet-Bahri", veladet);
  await save("Mevlid", "Mirac-Bahri", mirac);
  await save("Mevlid", "Rihlet-Bahri", rihlet);
  await save("Mevlid/Osmanlica", "0-Tanitim", ottTanitim);
  await save("Mevlid/Osmanlica", "Tevhid-Bahri", ottDoc("TEVHİD BAHRİ (OSMANLICA, HAREKELİ)", "Allah adın zikredelim evvelâ", ottTevhid));
  await save("Mevlid/Osmanlica", "Veladet-Bahri", ottDoc("VELÂDET BAHRİ (OSMANLICA, HAREKELİ)", "Âmine hâtun Muhammed anası", ottVeladet));
  await save("Mevlid/Osmanlica", "Merhaba-Bahri", ottDoc("MERHABA BAHRİ (OSMANLICA, HAREKELİ)", "Yâradılmış cümle oldu şâdümân", ottMerhaba));
  await save("Mevlid/Osmanlica", "Mirac-Bahri", ottDoc("MİRAÇ BAHRİ (OSMANLICA, HAREKELİ)", "Hak anı Mi'râca okudı gece", ottMirac));
  await save("Mevlid/Osmanlica", "Rihlet-Bahri", ottDoc("RİHLET BAHRİ (OSMANLICA, HAREKELİ)", "Şol rivâyet kim anuñ vefâtınuñ", ottRihlet));
  await save("Mevlid/Osmanlica", "Amine-Hatun-Bahri", ottDoc("ÂMİNE HÂTUN BAHRİ (OSMANLICA, HAREKELİ)", "Âmine hâtun Muhammed ânesi", ottAmine));
  await save("Mevlid/Osmanlica", "Dua-Iltica-Bahri", ottDoc("DUA VE İLTİCA BAHRİ (OSMANLICA, HAREKELİ)", "Yâ ilâhî sakla-gıl îmânımız", ottDuaIltica));
  await save("Salavatlar", "Salavatlar", salevat);
  await save("Ilahiler", "Mevlid_Ilahileri", ilahiler);
  await save("Kasideler", "Kasideler-Genel-Bilgi", kasideler);
  await save("Kasideler", "Kaside-i-Burde-Giris", burde);
  console.log("\nTamamlandi.");
})();
