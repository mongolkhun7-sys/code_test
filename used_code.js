// --- КОДЫН ЭХЛЭЛ ---

/****************************************************************************************
 * ЭНЭ КОД ЯМАР ЗОРИУЛАЛТТАЙ ВЭ?
 *
 * Энэхүү Google Apps Script (GAS) код нь харилцагчийн илгээсэн замбараагүй датаг
 * (жишээ нь: Хүйс, Нас, Өндөр, Жин) Google Sheet-ээс уншиж аваад,
 * Gemini AI ашиглан тэдгээрийг системтэй дата (JSON) болгон салгаж,
 * улмаар харилцагчийн биеийн жингийн индекс (БЖИ буюу BMI)-ийг тооцоолон,
 * Саруулбат Коучийн өнгө аястай, 7 хэсэгтэй хувийн тайланг AI-гаар бичүүлж,
 * Google Doc загвар ашиглан автоматаар PDF үүсгээд,
 * Uchat платформоор дамжуулан тухайн харилцагч руу илгээх зориулалттай систем юм.
 ****************************************************************************************/

const PRODUCT_CONFIG = {
  VERSION: "v25.2",
  COACH_NAME: "Халиунаа",
  PRODUCT_NAME: "Хувийн Жингийн Тайлан",
  SHEET_NAME: "Sheet1",
  SEND_ERROR_EMAILS: true,
  REPORT_DISCLAIMER_LINE: "Жич: Энэхүү тайлан нь зөвхөн ерөнхий эрүүл мэндийн боловсрол олгох зорилготой бөгөөд эмнэлгийн оношилгоо, эмчилгээг орлохгүй.",
  REPORT_SIGNATURE_LINE: "Чиний коуч: Халиунаа",

  COLUMNS: {
    NAME: 0, ID: 1, INPUT: 2, PDF: 3, STATUS: 4,
    TOKEN: 5, TYPE: 6, DATE: 7, VER: 8, ERROR: 9
  },

  UCHAT: {
    ENDPOINT: "https://www.uchat.com.au/api/subscriber/send-content",
    DELIVERY_MESSAGE: `Хөөх найз аа, {{NAME}}! 🎉\n\nЧиний хүлээж байсан, нүдийг чинь нээх нарийвчилсан тайланг чинь бэлдээд дуусчихлаа. 😉 \nХалиунаа нь байна аа. Доорх товч дээр дараад шууд татаад аваарай. 👇`,
    DELIVERY_BTN_TEXT: "🔥 Тайлангаа харъя"
  },

  BATCH_SIZE: 3,
  GEMINI_MODEL: "gemini-2.5-flash",
  TEMPERATURE: 0.4,
  REPORT_QUALITY: {
    SAFE_EMOJIS: ["👋", "🤔", "📈", "🚨", "😴", "🧠", "🔥", "💪", "🥗", "🍳", "🥣", "🍎", "🥩", "🍚", "🚶", "✅", "⚠", "🌿", "💧", "🚀", "📌", "✨", "🏋"]
  },

  PROMPTS: {
    EXTRACTOR: `
      You are a precise Data Extraction tool.
      Read the user's raw input below and extract key data into a strict JSON format.
      The expected format is usually: [Хүйс] - [Нас] - [Өндөр] - [Жин]
      Example: "ЭРЭГТЭЙ - 30 - 175 - 80" or "Эм - 165 - 65".

      RULES:
      1. Height is usually the number between 140 and 200 (cm).
      2. Weight is usually the number between 40 and 150 (kg).
      3. Age is usually between 15 and 70.
      4. IF age is missing (e.g. "Эр - 175 - 80"), DO NOT GUESS. Return null for age.
      5. GENDER: MUST be EXACTLY "эрэгтэй" or "эмэгтэй" in Cyrillic. NEVER output "male", "female", "man", or "woman" in English.

      RAW INPUT:
      "{{RAW_DATA}}"

      REQUIRED JSON FORMAT:
      {
        "age": Number or null,
        "gender": String (EXACTLY "эрэгтэй" OR "эмэгтэй"),
        "weight": Number (current weight in kg),
        "height": Number (cm)
      }
    `,

    SYSTEM_ROLE: `
      Таны нэр бол Халиунаа (Khaliunaa). Та бол маш "sassy", өөртөө итгэлтэй, ухаалаг бөгөөд хүмүүсийн хамгийн дотны найз шиг мөртлөө яг үнэнийг шууд хэлдэг коуч.

      >>> ⛔ ХАТУУ ДҮРМҮҮД (ШУУД ДАГАЖ МӨРДӨХ):
      1. **ХЭЛ ЗҮЙ БА ӨГҮҮЛБЭРИЙН БҮТЭЦ:** Монгол хэлний дүрмээр Эзэн бие - Тусагдахуун - Үйл үг (SOV) дарааллаар бичнэ. Англи хэлний (SVO) бүтцээр шууд үгчлэн орчуулахыг ХАТУУ ХОРИГЛОНО.
      2. **ШУУД ОРЧУУЛГЫН АЛДААГ ХОРИГЛОНО:**
         - "Sugar-free" = "сахаргүй" (хэзээ ч "элсгүй" гэж болохгүй).
         - "Warm-up" = "Халаалт" (хэзээ ч "Дулаалга" гэж болохгүй).
         - "Cool-down" = "Суллах дасгал" (хэзээ ч "Хөргөлт" гэж болохгүй).
         - "Plank" = "Планк" (хэзээ ч "Самбар" гэж болохгүй).
         - "Jumping Jacks" = "Үсрэлттэй дасгал" эсвэл шууд "Jumping Jacks".
         - Дасгалын нэрийг зохиож махчилан орчуулж инээдтэй болгож огт болохгүй! Шаардлагатай бол шууд Англиар нь хаалтанд бич (жишээ: Планк /Plank/).
      3. **ХЭСГҮҮДИЙН ЗАЛГААС (NO REPETITIVE INTROS):** 2 болон 3-р хэсгийг эхлэхдээ "Сайн уу", "Халиунаа байна", "Эрхэмээ" гэж ХЭЗЭЭ Ч БҮҮ МЭНДЭЛ! Шууд л өмнөх хэсгээсээ үргэлжилж байгаа мэтээр шууд сэдэв рүүгээ орно.
      4. **ХҮЙСИЙН МЭДРЭМЖ:** Одоо харьцаж байгаа хүн бол {{GENDER}}. Хүйсэнд тохирсон үг хэллэг ашигла.
	      5. **ГАРЧИГ БИЧИХ ДҮРЭМ (МАШ ЧУХАЛ):** Гарчиг бүрийнхээ урд ЗААВАЛ [ГАРЧИГ] гэдэг тагийг яг энэ чигээр нь бичнэ. Зөвхөн энэ үгийг харж код гарчгийг томруулдаг тул ОГТ МАРТАЖ БОЛОХГҮЙ. (Жишээ нь: "[ГАРЧИГ] Чиний биеийн нууц код"). Гарчиг дотор эможи БАЙЖ БОЛОХГҮЙ.
	      6. **ЭМОЖИ ДҮРЭМ (ОНЦГОЙ АНХААР):** ТАЙЛАНГИЙН БҮХ ХЭСГИЙН (ЭХЛЭЛЭЭС ТӨГСГӨЛ ХҮРТЭЛХ) ДОГОЛ МӨР БОЛОН ЖАГСААЛТ БҮРИЙН ХАМГИЙН ЭХЭНД ЗААВАЛ НЭГ ЭМОЖИ БАЙХ ЁСТОЙ! Өгүүлбэрийн дунд, эсвэл төгсгөлд эможи ОГТ БҮҮ ТАВЬ!
	      7. **ФОРМАТ:** Хэт их нуршиж давтсан үг БҮҮ ашигла. Яг гол санаагаа товч, тодорхой, хүчтэй илэрхийл.
	      8. **ЭРҮҮЛ МЭНДИЙН ХЯЗГААР:** Абсолют хориг, эмчилгээний заалт, эсвэл оношлох өнгө аястай тушаал БҮҮ өг. Зөвхөн ерөнхий эрүүл мэндийн боловсролын хүрээнд зөвлө. Хэрэв тухайн зөвлөгөө нь архаг өвчин, гэмтэл, жирэмслэлт, эм уудаг байдал, эсвэл эмчийн тусгай заавраас хамаарч өөрчлөгдөж болох бол заавал мэргэжлийн эмчтэй тулгаж хэрэгжүүлэхийг сануул.
	      9. **SAFE EMOJI ЖАГСААЛТ:** Зөвхөн дараах эможинуудыг ашигла: 👋 🤔 📈 🚨 😴 🧠 🔥 💪 🥗 🍳 🥣 🍎 🥩 🍚 🚶 ✅ ⚠ 🌿 💧 🚀 📌 ✨ 🏋. Эдгээрийг зөвхөн мөрийн эхэнд тавина.
	    `,



    FITNESS_OBESE: {
      PART_1: `
I. РОЛЬ: {{ROLE}}
II. МЭДЭЭЛЭЛ: Нэр: {{NAME}}, Нас: {{AGE}}, Хүйс: {{GENDER}}, Өндөр: {{HEIGHT}} см, Жин: {{WEIGHT}} кг, БЖИ: {{CALC_BMI}}, Илүүдэл: {{EXCESS_WEIGHT}} кг
III. ДААЛГАВАР: Эхний хэсэг (Бодит байдал ба Далд шалтгаанууд).
V. БҮТЭЦ:
   [ГАРЧИГ] Нүдээ нээ, найз минь! (Чиний бодит байдал)
      - Шууд энэ гарчгаар эхэлнэ. (Энд 1 удаа "За найз минь" гэх мэтээр эхэлж болно).
      - {{HEIGHT}} см өндөртэй хүнд {{WEIGHT}} кг жин гэдэг өвдөг, нуруу, зүрхэнд ямар их ачаалал болохыг маш товч бөгөөд хүчтэй хэл.
      - BMI ({{CALC_BMI}}) болон {{EXCESS_WEIGHT}} кг илүүдэл өөхний аюулыг сануул.
   [ГАРЧИГ] Чиний нууц дайснууд (Яагаад таргалаад байна вэ?)
      - 1. Инсулины дөжрөл: Байнга юм үмхлэх үед өөх шатаах хаалга яаж түгжигддэг тухай товчхон.
      - 2. Нойр ба Стресс: Оройтож унтах нь өлсгөлөнгийн даавар (Грелин)-ыг хэрхэн нэмэгдүүлдэг талаар товч хэл.
VI. ЗААВАР: Нийт 900-1000 орчим үг. Догол мөрний эхэнд ганц эможи байна. Гарчгийн урд ЗААВАЛ "[ГАРЧИГ]" гэдэг үг бичнэ.
`,
      PART_2: `
I. РОЛЬ: {{ROLE}}
II. МЭДЭЭЛЭЛ: Нас: {{AGE}}, Хүйс: {{GENDER}}, Жин: {{WEIGHT}} кг. Илүүдэл: {{EXCESS_WEIGHT}} кг.
III. ДААЛГАВАР: 2 дахь хэсэг (Худал төөрөгдөл ба Нарийвчилсан ХООЛНЫ төлөвлөгөө).
V. БҮТЭЦ:
   [ГАРЧИГ] Худлаа төөрөгдлүүдээс салцгаая!
      - АНХААР: "Сайн уу", "Халиунаа байна" гэж ОГТ МЭНДЛЭХГҮЙ, уулга алдахгүй ШУУД ҮРГЭЛЖЛҮҮЛ.
      - Зүгээр л хоолоо сойж өлсөх (Crash diet) яагаад чамайг буцаад таргалуулдгийг шүүмжил. Калорийн алдагдал болон уургийн дулааны нөлөө (TEF)-г тайлбарла.
   [ГАРЧИГ] Маргаашнаас юу өөрчлөгдөх вэ? (Mindset)
      - Дадал зуршлаа өөрчлөх тухай маш хүчтэй, богинохон урам зориг.
	   [ГАРЧИГ] Өлсөхгүйгээр турах: Хоолны төлөвлөгөө
	      - ЯГ ЭНЭ 3 ГАРЧГИЙГ ДАРААЛЛААР НЬ, ТУС БҮР НЭГ УДАА АШИГЛА. Аль нэгийг нь алгасаж БОЛОХГҮЙ.
	      - Энэ хэсэгт дасгалын хуваарь, дасгалын өдөр, сет/давталт ОГТ БИЧИХГҮЙ.
	      - Хоолоо идэх дараалал: 1. Ногоо -> 2. Уураг -> 3. Нүүрс ус (Маш товч).
	      - Өглөө: Уургаар баялаг цэсийг шууд жагсаа (сахаргүй тараг гэж бичнэ).
	      - Өдөр: Нүүрс ус, Уураг, Ногооны харьцааг заа.
	      - Орой: Унтахаас 2-3 цагийн өмнө хөнгөн, уурагтай хооллох ерөнхий зөвлөмж өг. Ходоод, чихрийн шижин, жирэмслэлт, эсвэл эмчийн заавартай бол хоолны цагийг мэргэжлийн хүнтэй тулгаж тохируулахыг сануул.
	VI. ЗААВАР: Нийт 900-1000 орчим үг. Хоолны хэсэг нь цэвэр жагсаалт байна. Гарчгийн урд ЗААВАЛ "[ГАРЧИГ]" гэдэг үг бичнэ.
	`,
      PART_3: `
I. РОЛЬ: {{ROLE}}
II. МЭДЭЭЛЭЛ: Нас: {{AGE}}, Хүйс: {{GENDER}}, Жин: {{WEIGHT}} кг.
III. ДААЛГАВАР: 3 дахь хэсэг (Нарийвчилсан ДАСГАЛЫН хуваарь ба төгсгөл).
V. БҮТЭЦ:
   [ГАРЧИГ] Өөхөө шатаах: Дасгалын хуваарь
      - АНХААР: "Сайн уу", "Халиунаа байна" гэж ОГТ МЭНДЛЭХГҮЙ ШУУД ҮРГЭЛЖЛҮҮЛ.
      - Нуршуу оршилгүйгээр шууд л хуваарь өг. (Халаалт, Суллах дасгал гэж бичнэ. "Дулаалга, Хөргөлт" ОГТ БИЧИХГҮЙ!)
	      - Энэ хэсэгт хоолны төлөвлөгөө, Өглөө/Өдөр/Орой цэс рүү БУЦАЖ ОРОХГҮЙ.
	      - Даваа (Хөл, өгзөг): (Жагсаалтаар бич).
	      - Лхагва (Дээд бие): (Жагсаалтаар бич).
	      - Баасан (Бүтэн бие): Zone 2 Cardio эсвэл хялбаршуулсан эрчимт дасгал (Жагсаалтаар бич. Планк гэж бичнэ, "Самбар" ОГТ БИЧИХГҮЙ).
	      - Бусад өдөр: Идэвхтэй амралт (Иого, сунгалт).
   [ГАРЧИГ] Эцсийн үг (Одоо чиний ээлж)
      - "Би чамд төлөвлөгөөг өглөө. Одоо хэрэгжүүлэх нь чамаас шалтгаална."
      - Заавал тайлангийн төгсгөлд дараах өгүүлбэрийг яг хуулж бич: "Жич: Энэхүү тайлан нь зөвхөн ерөнхий эрүүл мэндийн боловсрол олгох зорилготой бөгөөд эмнэлгийн оношилгоо, эмчилгээг орлохгүй."
      - Төгсгөлд нь "Чиний коуч: Халиунаа" гэж гарын үсэг зур. (Энд эможи тавьж болно).
VI. ЗААВАР: Нийт 700-800 орчим үг. Цэвэр жагсаалт байна. Гарчгийн урд ЗААВАЛ "[ГАРЧИГ]" гэдэг үг бичнэ.
`
    },
    FITNESS_NORMAL: {
      PART_1: `
I. РОЛЬ: {{ROLE}}
II. МЭДЭЭЛЭЛ: Нэр: {{NAME}}, Нас: {{AGE}}, Хүйс: {{GENDER}}, Өндөр: {{HEIGHT}} см, Жин: {{WEIGHT}} кг, БЖИ: {{CALC_BMI}}
III. ДААЛГАВАР: Эхний хэсэг (Баяр хүргэлт, Далд эрсдэлүүд).
V. БҮТЭЦ:
   [ГАРЧИГ] Баяр хүргэе, Найз аа! (Гэхдээ тайвшрах болоогүй)
      - Шууд энэ гарчгаар эхэлнэ. (Энд 1 удаа мэндэлж болно).
      - Жингээ хэвийн (BMI: {{CALC_BMI}}) барьж байгааг нь магт. Гэхдээ Skinny Fat (гаднаа туранхай ч дотроо өөхтэй) байхын аюулыг сануул.
   [ГАРЧИГ] Жин хэвийн ч гэсэн чиний анхаарах ёстой эрсдэлүүд
      - 1. Булчингийн алдагдал: Нас ахих тусам булчин хайлж өөх хуримтлагддаг тухай товч.
      - 2. Инсулины мэдрэмж: Туранхай байлаа гээд чихэр хамаагүй идэж болохгүйг анхааруул.
VI. ЗААВАР: Нийт 900-1000 орчим үг. Догол мөрний эхэнд ганц эможи байна. Гарчгийн урд ЗААВАЛ "[ГАРЧИГ]" гэдэг үг бичнэ.
`,
      PART_2: `
I. РОЛЬ: {{ROLE}}
II. МЭДЭЭЛЭЛ: Хэвийн жинтэй, Хүйс: {{GENDER}}.
III. ДААЛГАВАР: 2 дахь хэсэг (Галбиржих урлаг ба Нарийвчилсан ХООЛНЫ төлөвлөгөө).
V. БҮТЭЦ:
   [ГАРЧИГ] Галбиржих урлаг (Body Recomposition)
      - АНХААР: "Сайн уу", "Халиунаа байна" гэж ОГТ МЭНДЛЭХГҮЙ ШУУД ҮРГЭЛЖЛҮҮЛ. Хэвийн жинтэй хүний зорилго "турах" биш, харин "галбиржих" байх ёстойг хэл. Булчинд уураг чухал.
   [ГАРЧИГ] Ирээдүйдээ хийх хөрөнгө оруулалт (Mindset)
      - Өнөөдөр хийсэн хүчний дасгал чинь 10, 20 жилийн дараах залуу насыг хамгаалах "Anti-aging" нууц гэдгийг товч хэл.
   [ГАРЧИГ] Галбиржих Хоолны Төлөвлөгөө
	      - ЯГ ЭНЭ 3 ГАРЧГИЙГ ДАРААЛЛААР НЬ, ТУС БҮР НЭГ УДАА АШИГЛА. Аль нэгийг нь алгасаж БОЛОХГҮЙ.
	      - Энэ хэсэгт дасгалын хуваарь, дасгалын өдөр, сет/давталт ОГТ БИЧИХГҮЙ.
	      - Хооллолт: Зүгээр л бага идэх биш, Цусан дахь сахарыг тогтвортой барих нууц (Ногоо -> Уураг -> Нүүрс ус).
	      - Өглөө: Уураг болон эрүүл өөх тос цэсийг шууд жагсаа.
	      - Өдөр: Тэнцвэртэй харьцаа, булчин тэжээх уураг.
	      - Орой: Хөнгөн боловч булчин нөхөн төлжүүлэх уурагтай хоол.
VI. ЗААВАР: Нийт 900-1000 орчим үг. Хоолны хэсэг нь цэвэр жагсаалт байна. Гарчгийн урд ЗААВАЛ "[ГАРЧИГ]" гэдэг үг бичнэ.
`,
      PART_3: `
I. РОЛЬ: {{ROLE}}
II. МЭДЭЭЛЭЛ: Хэвийн жинтэй, Хүйс: {{GENDER}}.
III. ДААЛГАВАР: 3 дахь хэсэг (Нарийвчилсан ДАСГАЛЫН хуваарь ба төгсгөл).
V. БҮТЭЦ:
   [ГАРЧИГ] Булчин чангалах Дасгалын Хуваарь
      - АНХААР: "Сайн уу", "Халиунаа байна" гэж ОГТ МЭНДЛЭХГҮЙ ШУУД ҮРГЭЛЖЛҮҮЛ. Нуршуу оршилгүйгээр шууд л хуваарь өг. (Халаалт, Суллах дасгал гэж бичнэ. "Дулаалга, Хөргөлт" ОГТ БИЧИХГҮЙ!)
      - Унтах үед л булчин томорч, бие залууждаг (HGH даавар) тухай товч сануул (1 өгүүлбэр).
	      - Энэ хэсэгт хоолны төлөвлөгөө, Өглөө/Өдөр/Орой цэс рүү БУЦАЖ ОРОХГҮЙ.
	      - Даваа (Хөл, өгзөг): (Жагсаалтаар бич).
	      - Лхагва (Дээд бие): (Жагсаалтаар бич).
	      - Баасан (Бүтэн бие): Хүчний дасгал (Жагсаалтаар бич. Планк гэж бичнэ, "Самбар" ОГТ БИЧИХГҮЙ).
      - Бусад өдөр: Идэвхтэй амралт (Иого, сунгалт).
   [ГАРЧИГ] Эцсийн үг (Одоо чиний ээлж)
      - "Би чамд төлөвлөгөөг өглөө. Одоо хэрэгжүүлэх нь чамаас шалтгаална."
      - Заавал тайлангийн төгсгөлд дараах өгүүлбэрийг яг хуулж бич: "Жич: Энэхүү тайлан нь зөвхөн ерөнхий эрүүл мэндийн боловсрол олгох зорилготой бөгөөд эмнэлгийн оношилгоо, эмчилгээг орлохгүй."
      - Төгсгөлд нь "Чиний коуч: Халиунаа" гэж гарын үсэг зур.
VI. ЗААВАР: Нийт 700-800 орчим үг. Цэвэр жагсаалт байна. Гарчгийн урд ЗААВАЛ "[ГАРЧИГ]" гэдэг үг бичнэ.
`
    },
    FITNESS_UNDERWEIGHT: {
      PART_1: `
I. РОЛЬ: {{ROLE}}
II. МЭДЭЭЛЭЛ: Нэр: {{NAME}}, Нас: {{AGE}}, Хүйс: {{GENDER}}, Өндөр: {{HEIGHT}} см, Жин: {{WEIGHT}} кг, БЖИ: {{CALC_BMI}}
III. ДААЛГАВАР: Эхний хэсэг (Бодит байдал, Нуугдмал аюулууд).
V. БҮТЭЦ:
   [ГАРЧИГ] Үнэнтэй нүүр тулъя, найз минь!
      - Шууд энэ гарчгаар эхэлнэ. (Энд 1 удаа мэндэлж болно).
      - {{HEIGHT}} см өндөртэй хүнд {{WEIGHT}} кг жин гэдэг хэт туранхай буюу жингийн дутагдалтай байгааг анхааруул. Туранхай байх нь тарган байхаас дутахааргүй эрсдэлтэйг товч хэл.
   [ГАРЧИГ] Туранхай байхын нуугдмал аюулууд
      - 1. Дархлаа ба Шим тэжээлийн дутагдал: Биед өөх тос, уургийн нөөц байхгүй бол өвчин эсэргүүцэх чадвар унадаг тухай товч.
      - 2. Ясны хэврэгшил: Булчин байхгүй бол яс чинь хэврэг болдог талаар товч.
      - 3. Дааврын уналт: Нөхөн үржихүйн даавар алдагдах талаар товч.
VI. ЗААВАР: Нийт 900-1000 орчим үг. Догол мөрний эхэнд ганц эможи байна. Гарчгийн урд ЗААВАЛ "[ГАРЧИГ]" гэдэг үг бичнэ.
`,
      PART_2: `
I. РОЛЬ: {{ROLE}}
II. МЭДЭЭЛЭЛ: Жингийн дутагдалтай, Хүйс: {{GENDER}}.
III. ДААЛГАВАР: 2 дахь хэсэг (Булчин нэмэх шинжлэх ухаан ба ХООЛНЫ төлөвлөгөө).
V. БҮТЭЦ:
   [ГАРЧИГ] Булчин нэмэх шинжлэх ухаан (Зүгээр л таргалах биш!)
      - АНХААР: "Сайн уу", "Халиунаа байна" гэж ОГТ МЭНДЛЭХГҮЙ ШУУД ҮРГЭЛЖЛҮҮЛ. Калорийн илүүдэл (Caloric Surplus) үүсгэх хуулийг тайлбарла. Зүгээр л чихэр, өөх тос идэх биш Цэвэр булчин нэмэхийн чухлыг онцол.
   [ГАРЧИГ] Сэтгэлзүйн ялалт (Mindset)
      - "Би төрөлхийн туранхай" гэдэг шалтаг хэлэхээ зогсоо гэж загна. Хүн бүр булчин хөгжүүлж чадна гэдэгт итгүүл.
   [ГАРЧИГ] Эрүүл жин нэмэх Хоолны Төлөвлөгөө
	      - ЯГ ЭНЭ 3 ГАРЧГИЙГ ДАРААЛЛААР НЬ, ТУС БҮР НЭГ УДАА АШИГЛА. Аль нэгийг нь алгасаж БОЛОХГҮЙ.
	      - Энэ хэсэгт дасгалын хуваарь, дасгалын өдөр, сет/давталт ОГТ БИЧИХГҮЙ.
	      - Хоолны дуршил бага байсан ч калори ихтэй, тэжээллэг хүнс (самар, авокадо, бүтэн сүү) хэрхэн хоолондоо нууж шингээх аргууд.
	      - Өглөө: Калори өндөртэй, тэжээллэг хоолыг шууд жагсаа.
	      - Өдөр: Нүүрс ус (төмс, будаа) ихтэй, мэдээж уураг орсон цэсийг заа.
      - Орой: "Шингэн калори" (шэйк) ашиглах Халиунаагийн аргыг заа.
VI. ЗААВАР: Нийт 900-1000 орчим үг. Хоолны хэсэг нь цэвэр жагсаалт байна. Гарчгийн урд ЗААВАЛ "[ГАРЧИГ]" гэдэг үг бичнэ.
`,
      PART_3: `
I. РОЛЬ: {{ROLE}}
II. МЭДЭЭЛЭЛ: Жингийн дутагдалтай, Хүйс: {{GENDER}}.
III. ДААЛГАВАР: 3 дахь хэсэг (Нарийвчилсан ДАСГАЛЫН хуваарь ба төгсгөл).
V. БҮТЭЦ:
		   [ГАРЧИГ] Булчингаар томрох Дасгалын Хуваарь
		      - АНХААР: "Сайн уу", "Халиунаа байна" гэж ОГТ МЭНДЛЭХГҮЙ ШУУД ҮРГЭЛЖЛҮҮЛ. Нуршуу оршилгүйгээр шууд л хуваарь өг. (Халаалт, Суллах дасгал гэж бичнэ. "Дулаалга, Хөргөлт" ОГТ БИЧИХГҮЙ!)
		      - Унтах үед л бие булчингаа томруулж засварладаг гэдгийг 1 өгүүлбэрээр сануул.
		      - Энэ хэсэгт хоолны төлөвлөгөө, Өглөө/Өдөр/Орой цэс рүү БУЦАЖ ОРОХГҮЙ.
		      - Даваа (Хөл, өгзөг): (Жагсаалтаар бич).
		      - Лхагва (Цээж, гар): Түлхэлт, таталт (Жагсаалтаар бич).
		      - Баасан (Нуруу, Core): (Жагсаалтаар бич. Планк гэж бичнэ, "Самбар" ОГТ БИЧИХГҮЙ).
	      - Хэт удаан, хэт ачаалалтай кардиог багасгаж, хүчний бэлтгэлдээ төвлөрөх ерөнхий зөвлөмж өг. Амьсгал, зүрх судас, гэмтэлтэй бол дасгалаа эмч эсвэл мэргэжлийн багштай тохируулахыг сануул.
	   [ГАРЧИГ] Эцсийн үг (Одоо чиний ээлж)
	      - "Би чамд төлөвлөгөөг өглөө. Одоо хэрэгжүүлэх нь чамаас шалтгаална."
      - Заавал тайлангийн төгсгөлд дараах өгүүлбэрийг яг хуулж бич: "Жич: Энэхүү тайлан нь зөвхөн ерөнхий эрүүл мэндийн боловсрол олгох зорилготой бөгөөд эмнэлгийн оношилгоо, эмчилгээг орлохгүй."
      - Төгсгөлд нь "Чиний коуч: Халиунаа" гэж гарын үсэг зур.
VI. ЗААВАР: Нийт 700-800 орчим үг. Цэвэр жагсаалт байна. Гарчгийн урд ЗААВАЛ "[ГАРЧИГ]" гэдэг үг бичнэ.
`
    }
  }
};

function getProperty(key) {
  const val = PropertiesService.getScriptProperties().getProperty(key);
  if (!val) throw new Error(`MISSING SCRIPT PROPERTY: ${key}`);
  return val;
}

function main() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCT_CONFIG.SHEET_NAME);
    const rows = sheet.getDataRange().getValues();
    let processedCount = 0;

    const KEYS = {
      GEMINI: getProperty("GEMINI_API_KEY"),
      TEMPLATE: getProperty("TEMPLATE_ID"),
      UCHAT: getProperty("UCHAT_API_KEY"),
      FOLDER: getProperty("FOLDER_ID")
    };

    const COLS = PRODUCT_CONFIG.COLUMNS;

    for (let i = 1; i < rows.length; i++) {
      if (processedCount >= PRODUCT_CONFIG.BATCH_SIZE) break;

      const row = rows[i];
      const name = row[COLS.NAME];
      const contactID = row[COLS.ID];
      const inputData = String(row[COLS.INPUT]);
      const status = String(row[COLS.STATUS] || "");
      const rawDate = row[COLS.DATE];

      const pdfCell = sheet.getRange(i + 1, COLS.PDF + 1);
      const statusCell = sheet.getRange(i + 1, COLS.STATUS + 1);
      const errorCell = sheet.getRange(i + 1, COLS.ERROR + 1);
      const tokenCell = sheet.getRange(i + 1, COLS.TOKEN + 1);
      const typeCell = sheet.getRange(i + 1, COLS.TYPE + 1);
      const dateCell = sheet.getRange(i + 1, COLS.DATE + 1);
      const verCell = sheet.getRange(i + 1, COLS.VER + 1);

      if (!name || !inputData) continue;
      if (status === "АМЖИЛТТАЙ" || status.includes("ХЯНАХ ШААРДЛАГАТАЙ") || status.includes("24 цаг хэтэрсэн")) continue;

      let isRetry = false;
      if (status === "Боловсруулж байна...") {
        let diffMinutes = 0;
        let validDate = false;

        if (rawDate instanceof Date) {
            diffMinutes = (new Date().getTime() - rawDate.getTime()) / (1000 * 60);
            validDate = true;
        } else if (typeof rawDate === "string" && rawDate.length > 5) {
            const parsedDate = new Date(rawDate);
            if (!isNaN(parsedDate.getTime())) {
                diffMinutes = (new Date().getTime() - parsedDate.getTime()) / (1000 * 60);
                validDate = true;
            }
        }

        if (validDate) {
            if (diffMinutes > 15) {
                isRetry = true;
                console.log(`Timeout recovery for ${name}. Stuck for ${Math.round(diffMinutes)} mins.`);
            } else {
                continue; // Still processing, let it be
            }
        } else {
            // If date is completely invalid/missing but stuck in Processing, force retry
            isRetry = true;
        }
      }

      statusCell.setValue("Боловсруулж байна...");

      const startTime = new Date();
      // Only set to formatted string to keep sheet clean, logic above now parses it.
      dateCell.setValue(Utilities.formatDate(startTime, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"));
      SpreadsheetApp.flush();

	      let pdfAsset = null;
	      let deliverySent = false;
	      try {
	        console.log(`Processing user: ${name}`);
	        const fullName = String(name);
	        const firstNameOnly = fixMongolianName(fullName.split(" ")[0]);

        const { data: cleanData, usage: extractUsage } = extractDataWithAI(inputData, KEYS.GEMINI);

        const metrics = calculateMetricsFromJSON(cleanData);

        let reportResult = generateReport3Parts(firstNameOnly, metrics, KEYS.GEMINI);
        let reportText = reportResult.text;
        let totalTokenUsage = extractUsage + reportResult.usage;

	        const existingPdfUrl = String(pdfCell.getValue() || "").trim();
	        if (existingPdfUrl) {
	          pdfAsset = { url: existingPdfUrl, fileId: null, reused: true };
	        } else {
	          pdfAsset = createPdfFromTemplate(fullName, reportText, KEYS.TEMPLATE, KEYS.FOLDER);
	          pdfCell.setValue(pdfAsset.url);
	          SpreadsheetApp.flush();
	        }

	        sendUChatProven(contactID, pdfAsset.url, firstNameOnly, KEYS.UCHAT);
	        deliverySent = true;

	        const now = new Date();
	        const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

	        pdfCell.setValue(pdfAsset.url);
	        tokenCell.setValue(totalTokenUsage);
	        typeCell.setValue(metrics.BMI_CATEGORY);
	        dateCell.setValue(formattedDate);
        verCell.setValue(PRODUCT_CONFIG.VERSION);

        statusCell.setValue("АМЖИЛТТАЙ");
        errorCell.setValue("");

        processedCount++;

	      } catch (err) {
	        let errorMsgStr = err.toString();
	        let mongolianError = "Системийн алдаа гарлаа: " + errorMsgStr;
	        let requiresManualReview = false;

	        if (errorMsgStr.includes("24H_LIMIT") || errorMsgStr.includes("window")) {
	            if (pdfAsset && !deliverySent && !pdfAsset.reused) {
	              trashDriveFileQuietly(pdfAsset.fileId);
	              pdfCell.setValue("");
	            }
	            mongolianError = "Фэйсбүүк 24 цаг хэтэрсэн тул мессеж явуулах эрх хаагдсан байна.";
	            statusCell.setValue("24 цаг хэтэрсэн");
	            errorCell.setValue(mongolianError);
	            console.warn(`Skipped ${name}: 24h limit.`);
	            continue;
	        } else if (errorMsgStr.includes("DATA_EXTRACTION_PARSE_ERROR") || errorMsgStr.includes("DATA_VALIDATION_ERROR")) {
	            mongolianError = "Оролтын өгөгдөл дутуу эсвэл ойлгомжгүй тул гараар шалгах шаардлагатай.";
	            requiresManualReview = true;
	        } else if (errorMsgStr.includes("REPORT_STRUCTURE_ERROR")) {
	            mongolianError = "Тайлангийн бүтэц эвдэрсэн эсвэл гарчиг дутуу гарсан тул дахин оролдож байна.";
	        } else if (errorMsgStr.includes("Gemini") || errorMsgStr.includes("JSON Parse Error")) {
	            mongolianError = "AI (Gemini) хариу өгсөнгүй эсвэл түр зуур хэт ачаалалтай байна.";
	        } else if (errorMsgStr.includes("UChat token") || errorMsgStr.includes("user_ns")) {
	            mongolianError = "UChat тохиргоо эсвэл харилцагчийн код буруу байна.";
	        }

	        if (pdfAsset && !deliverySent && !pdfAsset.reused) {
	          trashDriveFileQuietly(pdfAsset.fileId);
	          pdfCell.setValue("");
	        }

	        console.error(`Error for ${name}: ${errorMsgStr}`);
	        errorCell.setValue(mongolianError);

	        if (requiresManualReview) {
	             statusCell.setValue("ХЯНАХ ШААРДЛАГАТАЙ");
	             sendErrorEmail(name, mongolianError);
	        } else if (isRetry || status === "") {
	             statusCell.setValue("Дахин оролдож байна (1)");
	        } else if (status === "Дахин оролдож байна (1)") {
	             statusCell.setValue("Дахин оролдож байна (2)");
        } else if (status === "Дахин оролдож байна (2)") {
             statusCell.setValue("ХЯНАХ ШААРДЛАГАТАЙ");
             sendErrorEmail(name, mongolianError);
        } else {
             statusCell.setValue("Дахин оролдож байна (1)");
        }
      }
    }
  } catch (e) {
    sendErrorEmail("SYSTEM_CRITICAL", e.toString());
  } finally {
    lock.releaseLock();
  }
}

function extractDataWithAI(rawInput, apiKey) {
  const prompt = PRODUCT_CONFIG.PROMPTS.EXTRACTOR.replace("{{RAW_DATA}}", rawInput);

  const result = callGeminiAPI(prompt, apiKey, 0.1, true);

  try {
    const data = JSON.parse(result.text.trim());
    return { data: data, usage: result.usage };
  } catch (e) {
    console.error("JSON Parse Error:", result.text);
    throw new Error(`DATA_EXTRACTION_PARSE_ERROR: ${e.message}`);
  }
}

function calculateMetricsFromJSON(data) {
  const age = parseOptionalNumber(data.age, "age", 15, 70);
  const weight = parseRequiredNumber(data.weight, "weight", 40, 150);
  const height = parseRequiredNumber(data.height, "height", 140, 200);

  // --- GENDER FIX (Англиар бичсэн ч монгол руу хөрвүүлнэ) ---
  let rawGender = (data.gender || "").toLowerCase().trim();
  let gender = null;

  if (rawGender === "эрэгтэй" || rawGender === "male" || rawGender === "man" || rawGender === "эр") {
      gender = "эрэгтэй";
  } else if (rawGender === "эмэгтэй" || rawGender === "female" || rawGender === "woman" || rawGender === "эм") {
      gender = "эмэгтэй";
  }

  if (!gender) {
      throw new Error("DATA_VALIDATION_ERROR: gender is missing or invalid.");
  }

  const heightM = height / 100;
  const bmi = weight / (heightM * heightM);
  const bmiFixed = bmi.toFixed(1);

  let category = "хэвийн жин";
  if (bmi < 18.5) category = "тураал";
  else if (bmi >= 25 && bmi < 30) category = "илүүдэл жин";
  else if (bmi >= 30) category = "таргалалт";

  const healthyWeightMax = 24.9 * (heightM * heightM);
  let excessWeight = 0;

  if (weight > healthyWeightMax) {
      excessWeight = Math.ceil(weight - healthyWeightMax);
      if (excessWeight < 3) excessWeight = "3-5";
      else excessWeight = `${excessWeight}-${excessWeight + 2}`;
  } else {
      excessWeight = "0-2";
  }

  return {
    AGE: age,
    GENDER: gender,
    WEIGHT: weight,
    HEIGHT: height,
    HEIGHT_M: heightM.toFixed(2),
    CALC_BMI: bmiFixed,
    BMI_CATEGORY: category,
    EXCESS_WEIGHT: excessWeight
  };
}

function generateReport3Parts(userName, metrics, apiKey) {
  const reportSpec = getReportSpec(metrics);
  const templateSet = reportSpec.templateSet;

  const variables = {
    "{{ROLE}}": PRODUCT_CONFIG.PROMPTS.SYSTEM_ROLE,
    "{{NAME}}": userName,
    "{{AGE}}": metrics.AGE !== null ? metrics.AGE : "null",
    "{{GENDER}}": metrics.GENDER,
    "{{WEIGHT}}": metrics.WEIGHT,
    "{{HEIGHT}}": metrics.HEIGHT,
    "{{HEIGHT_M}}": metrics.HEIGHT_M,
    "{{CALC_BMI}}": metrics.CALC_BMI,
    "{{BMI_CATEGORY}}": metrics.BMI_CATEGORY,
    "{{EXCESS_WEIGHT}}": metrics.EXCESS_WEIGHT
  };

  const replaceVars = (template) => {
    let result = template;
    result = result.split("{{ROLE}}").join(variables["{{ROLE}}"]);
    for (const [key, value] of Object.entries(variables)) {
      if (key !== "{{ROLE}}") {
        result = result.split(key).join(value);
      }
    }
    return result;
  };

  const withPreviousSections = (template, previousSections) => {
    if (!previousSections || previousSections.length === 0) return template;
    // VERY IMPORTANT FIX: Passing the entire previous text causes context pollution and duplicated headings/cut-offs.
    // Instead of passing the massive text back, just give the LLM a short summary of what was already covered.
    let contextSummary = "";
    const isObese = reportSpec.categoryLabel === "Таргалалттай тайлан";
    const isUnder = reportSpec.categoryLabel === "Жингийн дутагдалтай тайлан";

    if (previousSections.length === 1) {
        if (isObese) {
            contextSummary = "- Та өмнөх (1-р) хэсэгт бодит байдал, таргалалтын шалтгааныг (Инсулин, Нойр) тайлбарлаж дууссан.";
        } else if (isUnder) {
            contextSummary = "- Та өмнөх (1-р) хэсэгт бодит байдал, жингийн дутагдлын аюулыг тайлбарлаж дууссан.";
        } else {
            contextSummary = "- Та өмнөх (1-р) хэсэгт хэвийн болон илүүдэл жингийн далд эрсдэлүүдийг тайлбарлаж дууссан.";
        }
    } else if (previousSections.length === 2) {
        contextSummary = "- Та өмнөх (2-р) хэсэгт хоолны төлөвлөгөөг (Өглөө, Өдөр, Орой) зааж өгөөд дууссан. Тиймээс ХООЛНЫ ТАЛААР ДАХИЖ ДАВТАХГҮЙ шууд дасгалын хуваарь руу орно.";
    }

    return `${template}

VII. ӨМНӨХ ХЭСГҮҮДИЙН КОНТЕКСТ БА ҮРГЭЛЖЛЭЛИЙН ДҮРЭМ:
${contextSummary}
- Өмнөх хэсгүүдийн гарчгийг дахин ДАВТАХГҮЙ.
- Шинэ хэсгээ өмнөхөөсөө байгалийн байдлаар, зөвхөн шинэ мэдээллээр үргэлжлүүл.
- ХЭЗЭЭ Ч БҮҮ МЭНДЭЛ (Сайн уу гэх мэт).
- Текстээ бүрэн гүйцэд дуусгахыг хатуу анхаар! (Cut-off хийж болохгүй).`;
  };

  const p1 = replaceVars(templateSet.PART_1);
  const r1 = callGeminiAPI(p1, apiKey, PRODUCT_CONFIG.TEMPERATURE);

  const p2 = withPreviousSections(replaceVars(templateSet.PART_2), [true]); // Use dummy array to indicate part 2
  const r2 = callGeminiAPI(p2, apiKey, PRODUCT_CONFIG.TEMPERATURE);

  const p3 = withPreviousSections(replaceVars(templateSet.PART_3), [true, true]); // Use dummy array to indicate part 3
  const r3 = callGeminiAPI(p3, apiKey, PRODUCT_CONFIG.TEMPERATURE);

  const draftReport = [r1.text.trim(), r2.text.trim(), r3.text.trim()].join("\n\n");
  const sanitizedText = sanitizeReportText(draftReport, reportSpec);

  return {
    text: sanitizedText,
    usage: (r1.usage || 0) + (r2.usage || 0) + (r3.usage || 0)
  };
}

function getReportSpec(metrics) {
  const sharedPhrases = [
    "Өглөө:",
    "Өдөр:",
    "Орой:",
    "Халаалт",
    "Даваа",
    "Лхагва",
    "Баасан",
    "Бусад өдөр",
    "Суллах дасгал",
    PRODUCT_CONFIG.REPORT_DISCLAIMER_LINE,
    PRODUCT_CONFIG.REPORT_SIGNATURE_LINE
  ];

  if (metrics.BMI_CATEGORY === "тураал") {
    return {
      templateSet: PRODUCT_CONFIG.PROMPTS.FITNESS_UNDERWEIGHT,
      requiredHeadings: [
        "Үнэнтэй нүүр тулъя, найз минь!",
        "Туранхай байхын нуугдмал аюулууд",
        "Булчин нэмэх шинжлэх ухаан (Зүгээр л таргалах биш!)",
        "Сэтгэлзүйн ялалт (Mindset)",
        "Эрүүл жин нэмэх Хоолны Төлөвлөгөө",
        "Булчингаар томрох Дасгалын Хуваарь",
        "Эцсийн үг (Одоо чиний ээлж)"
      ],
      requiredPhrases: sharedPhrases,
      categoryLabel: "Жингийн дутагдалтай тайлан"
    };
  }

  if (metrics.BMI_CATEGORY === "хэвийн жин") {
    return {
      templateSet: PRODUCT_CONFIG.PROMPTS.FITNESS_NORMAL,
      requiredHeadings: [
        "Баяр хүргэе, Найз аа! (Гэхдээ тайвшрах болоогүй)",
        "Жин хэвийн ч гэсэн чиний анхаарах ёстой эрсдэлүүд",
        "Галбиржих урлаг (Body Recomposition)",
        "Ирээдүйдээ хийх хөрөнгө оруулалт (Mindset)",
        "Галбиржих Хоолны Төлөвлөгөө",
        "Булчин чангалах Дасгалын Хуваарь",
        "Эцсийн үг (Одоо чиний ээлж)"
      ],
      requiredPhrases: sharedPhrases,
      categoryLabel: "Хэвийн жингийн тайлан"
    };
  }

  return {
    templateSet: PRODUCT_CONFIG.PROMPTS.FITNESS_OBESE,
    requiredHeadings: [
      "Нүдээ нээ, найз минь! (Чиний бодит байдал)",
      "Чиний нууц дайснууд (Яагаад таргалаад байна вэ?)",
      "Худлаа төөрөгдлүүдээс салцгаая!",
      "Маргаашнаас юу өөрчлөгдөх вэ? (Mindset)",
      "Өлсөхгүйгээр турах: Хоолны төлөвлөгөө",
      "Өөхөө шатаах: Дасгалын хуваарь",
      "Эцсийн үг (Одоо чиний ээлж)"
    ],
    requiredPhrases: sharedPhrases,
    categoryLabel: metrics.BMI_CATEGORY === "илүүдэл жин" ? "Илүүдэл жингийн тайлан (Анхааруулах шат)" : "Таргалалттай тайлан"
  };
}

function sanitizeReportText(text, reportSpec) {
  let working = String(text || "")
    .replace(/\r\n?/g, "\n")
    .replace(/[\u200D\uFE0F\u2640\u2642]/g, "")
    .replace(/\t/g, " ")
    .replace(/\*\*/g, "")
    .replace(/^#+\s*/gm, "");

  const cleanedLines = [];
  const rawLines = working.split("\n");

  // Track seen headings to strictly remove duplicates
  const seenHeadings = new Set();

  for (const rawLine of rawLines) {
    const cleanLine = sanitizeReportLine(rawLine, reportSpec);
    if (!cleanLine) {
      if (cleanedLines.length > 0 && cleanedLines[cleanedLines.length - 1] !== "") {
        cleanedLines.push("");
      }
      continue;
    }

    if (cleanLine.startsWith("[ГАРЧИГ]")) {
        const headingName = cleanLine.replace("[ГАРЧИГ]", "").trim();
        if (seenHeadings.has(headingName)) {
            continue; // Skip duplicate headings entirely
        }
        seenHeadings.add(headingName);
    }

    cleanedLines.push(cleanLine);
  }

  working = cleanedLines.join("\n").replace(/\n{3,}/g, "\n\n").trim();

  // STRICT VALIDATION: Ensure all required headings are present and in the EXACT order
  let lastHeadingIndex = -1;
  let foundHeadingsCount = 0;

  for (const requiredHeading of reportSpec.requiredHeadings) {
      const currentHeadingIndex = working.indexOf(`[ГАРЧИГ] ${requiredHeading}`);
      if (currentHeadingIndex === -1) {
          throw new Error(`REPORT_STRUCTURE_ERROR: Missing required heading: ${requiredHeading}`);
      }
      if (currentHeadingIndex < lastHeadingIndex) {
          throw new Error(`REPORT_STRUCTURE_ERROR: Headings are out of order. '${requiredHeading}' appeared before its proper place.`);
      }
      lastHeadingIndex = currentHeadingIndex;
      foundHeadingsCount++;
  }

  if (foundHeadingsCount !== reportSpec.requiredHeadings.length) {
      throw new Error(`REPORT_STRUCTURE_ERROR: Incorrect number of headings found.`);
  }

  if (!working.includes(PRODUCT_CONFIG.REPORT_DISCLAIMER_LINE)) {
    working += `\n⚠ ${PRODUCT_CONFIG.REPORT_DISCLAIMER_LINE}`;
  }

  if (!working.includes(PRODUCT_CONFIG.REPORT_SIGNATURE_LINE)) {
    working += `\n✨ ${PRODUCT_CONFIG.REPORT_SIGNATURE_LINE}`;
  }

  return working.trim();
}

function sanitizeReportLine(line, reportSpec) {
  let working = String(line || "").replace(/\s+/g, " ").trim();
  if (!working) return "";
  if (/^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s/.test(working)) return "";

  const headingMatch = findHeadingMatch(working, reportSpec.requiredHeadings);
  if (headingMatch) {
    return `[ГАРЧИГ] ${headingMatch}`;
  }

  // --- AUTO-HEALING EMOJI LOGIC ---
  // Find if there is any emoji at the very beginning of the string
  const firstCharMatch = working.match(/^[\u{1F300}-\u{1F6FF}\u{1F900}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F1E6}-\u{1F1FF}\u{1F200}-\u{1F2FF}]/u);
  let startingEmoji = firstCharMatch ? firstCharMatch[0] : "";

  // Strip ALL emojis and markdown list bullets from the rest of the string
  working = working.replace(/[\u{1F300}-\u{1F6FF}\u{1F900}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F1E6}-\u{1F1FF}\u{1F200}-\u{1F2FF}]/gu, "")
                   .replace(/^[-*•]+\s*/, "")
                   .trim();

  if (!working) return "";

  const secondHeadingMatch = findHeadingMatch(working, reportSpec.requiredHeadings);
  if (secondHeadingMatch) {
    return `[ГАРЧИГ] ${secondHeadingMatch}`;
  }

  // Assign a relevant emoji if the AI didn't provide one at the start
  if (!startingEmoji) {
    startingEmoji = pickLineEmoji(working);
  }

  return `${startingEmoji} ${working}`;
}

function findHeadingMatch(line, requiredHeadings) {
  const normalizedLine = normalizeComparableText(line);

  // STRICT matching for headings to prevent false positives from normal sentences.
  // The line must almost exactly match the heading.
  for (const heading of requiredHeadings) {
    const normalizedHeading = normalizeComparableText(heading);

    // Direct exact match
    if (normalizedLine === normalizedHeading) {
        return heading;
    }

    // Allow slight variations (e.g., if AI misses a single short word or adds an extra space),
    // but the line cannot be a long paragraph containing the heading.
    if (normalizedLine.includes(normalizedHeading) && normalizedLine.length <= normalizedHeading.length + 10) {
        return heading;
    }
  }
  return null;
}

function normalizeComparableText(text) {
  return String(text || "")
    .replace(/[\u{1F300}-\u{1F6FF}\u{1F900}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F1E6}-\u{1F1FF}\u{1F200}-\u{1F2FF}]/gu, "")
    .replace(/\[ГАРЧИГ\]/g, "")
    .replace(/[\u200D\uFE0F\u2640\u2642()]/g, "")
    .replace(/^[-*•\d]+\.?\s*/, "")
    .replace(/[?!.]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function pickLineEmoji(line) {
  const text = String(line || "");

  if (/^(Өглөө:)/i.test(text)) return "🍳";
  if (/^(Өдөр:)/i.test(text)) return "🥗";
  if (/^(Орой:)/i.test(text)) return "🥣";
  if (/^(Халаалт)/i.test(text)) return "🔥";
  if (/^(Суллах дасгал)/i.test(text)) return "🌿";
  if (/^(Даваа|Лхагва|Баасан)/i.test(text)) return "🏋";
  if (/^(Бусад өдөр)/i.test(text)) return "🚶";
  if (/^(Жич:|ЭРҮҮЛ МЭНДИЙН АНХААРУУЛГА)/i.test(text)) return "⚠";
  if (/^(Чиний коуч:)/i.test(text)) return "✨";
  if (/^\d+\./.test(text)) return "📌";
  if (/(хоол|цэс|сахаргүй|тараг|шэйк|уураг|ногоо)/i.test(text)) return "🥗";
  if (/(дасгал|сет|давталт|планк|сквот|пресс|алхалт|кардио)/i.test(text)) return "💪";
  if (/(нойр|стресс|итгэл|дадал|сэтгэлзүй|mindset)/i.test(text)) return "🧠";
  if (/(жин|бми|bmi|эрсдэл|таргал|тураал)/i.test(text)) return "📈";
  return "✅";
}


function countMatches(text, regex) {
  const matches = String(text || "").match(regex);
  return matches ? matches.length : 0;
}

function callGeminiAPI(prompt, apiKey, temp, requireJson = false) {
  const model = PRODUCT_CONFIG.GEMINI_MODEL;
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  let payload = {
    "contents": [{ "parts": [{ "text": prompt }] }],
    "generationConfig": { "temperature": temp, "maxOutputTokens": 8192 },
    // --- SAFETY FILTER BYPASS ---
    "safetySettings": [
        { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_ONLY_HIGH" },
        { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_ONLY_HIGH" },
        { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_ONLY_HIGH" },
        { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_ONLY_HIGH" }
    ]
  };

  if (requireJson) {
      payload.generationConfig.responseMimeType = "application/json";
  }

  const res = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
  const json = JSON.parse(res.getContentText());

  if (json.candidates && json.candidates[0].content) {
      const candidate = json.candidates[0];
      const finishReason = candidate.finishReason || "";

      if (finishReason !== "STOP" && finishReason !== "") {
          throw new Error(`Gemini API Error: Abnormal finish reason - ${finishReason}. Response may be incomplete or filtered.`);
      }

      return {
          text: candidate.content.parts[0].text,
          usage: json.usageMetadata ? json.usageMetadata.totalTokenCount : 0
      };
  }

  throw new Error(`Gemini API Error: ${res.getContentText()}`);
}

function fixMongolianName(latinName) {
  if (!latinName || latinName.length === 0) return "Найз аа";
  let name = latinName.toLowerCase().trim();

  name = name.replace(/munkh/g, "мөнх").replace(/sukh/g, "сүх")
             .replace(/bat/g, "бат").replace(/erdene/g, "эрдэнэ")
             .replace(/bold/g, "болд").replace(/tulga/g, "тулга")
             .replace(/bayar/g, "баяр").replace(/naran/g, "наран");

  return name.split('-').map(part => part.charAt(0).toUpperCase() + part.slice(1)).join('-');
}

function parseRequiredNumber(value, fieldName, minValue, maxValue) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    throw new Error(`DATA_VALIDATION_ERROR: ${fieldName} is missing or invalid.`);
  }
  if (minValue !== undefined && parsed < minValue) {
    throw new Error(`DATA_VALIDATION_ERROR: ${fieldName} is below allowed range.`);
  }
  if (maxValue !== undefined && parsed > maxValue) {
    throw new Error(`DATA_VALIDATION_ERROR: ${fieldName} is above allowed range.`);
  }
  return parsed;
}

function parseOptionalNumber(value, fieldName, minValue, maxValue) {
  if (value === null || value === undefined || value === "") return null;
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    throw new Error(`DATA_VALIDATION_ERROR: ${fieldName} is invalid.`);
  }
  if (minValue !== undefined && parsed < minValue) {
    throw new Error(`DATA_VALIDATION_ERROR: ${fieldName} is below allowed range.`);
  }
  if (maxValue !== undefined && parsed > maxValue) {
    throw new Error(`DATA_VALIDATION_ERROR: ${fieldName} is above allowed range.`);
  }
  return parsed;
}

function isStructuredReportLine(text) {
  return /^(?:\S+\s+)?(Өглөө:|Өдөр:|Орой:|Халаалт|Суллах дасгал|Даваа|Лхагва|Баасан|Бусад өдөр|\d+\.)/i.test(text);
}

function getStructuredLeadEnd(text) {
  const colonIndex = text.indexOf(":");
  if (colonIndex > 0 && colonIndex < 45) return colonIndex;

  const numberedMatch = text.match(/^\S+\s+\d+\./);
  if (numberedMatch) return numberedMatch[0].length - 1;

  return -1;
}

function applyPremiumParagraphFormatting(paragraph, textObj, paragraphText, isHeading, inheritAttrs) {
  const baseFontSize = inheritAttrs ? inheritAttrs.fontSize : 11;
  const isStructuredLine = isStructuredReportLine(paragraphText);

  if (inheritAttrs) {
    textObj.setFontFamily(inheritAttrs.fontFamily);
    textObj.setForegroundColor(inheritAttrs.foregroundColor);
    textObj.setItalic(inheritAttrs.isItalic);
    textObj.setFontSize(isHeading ? baseFontSize + 2 : baseFontSize);
    textObj.setBold(isHeading ? true : inheritAttrs.isBold);
  }

  if (isHeading) {
    paragraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    paragraph.setSpacingBefore(16);
    paragraph.setSpacingAfter(8);
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    return;
  }

  paragraph.setLineSpacing(1.35);

  if (isStructuredLine) {
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    paragraph.setIndentStart(18);
    paragraph.setIndentFirstLine(0);
    paragraph.setSpacingBefore(4);
    paragraph.setSpacingAfter(6);

    const leadEnd = getStructuredLeadEnd(paragraphText);
    if (leadEnd >= 0) {
      textObj.setBold(0, leadEnd, true);
    }
    return;
  }

  if (paragraphText.length > 70) {
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  } else {
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
  }

  paragraph.setSpacingAfter(10);
}

function createPdfFromTemplate(name, content, templateId, folderId) {
  const targetFolder = DriveApp.getFolderById(folderId);
  const copyFile = DriveApp.getFileById(templateId).makeCopy(`${name} - ${PRODUCT_CONFIG.PRODUCT_NAME}`, targetFolder);
  const copyId = copyFile.getId();

  const doc = DocumentApp.openById(copyId);
  const body = doc.getBody();

  body.replaceText("{{name}}", name);
  body.replaceText("{{NAME}}", name);

  // --- FIND {{report}} AND INHERIT FORMATTING ---
  let reportElement = body.findText("{{report}}");
  let inheritAttrs = null;

  if (reportElement) {
      let textObj = reportElement.getElement().asText();
      let offset = reportElement.getStartOffset();

      inheritAttrs = {
          fontFamily: textObj.getFontFamily(offset),
          fontSize: textObj.getFontSize(offset) || 11, // fallback
          foregroundColor: textObj.getForegroundColor(offset),
          isBold: textObj.isBold(offset),
          isItalic: textObj.isItalic(offset)
      };

      // Remove {{report}} placeholder
      textObj.deleteText(reportElement.getStartOffset(), reportElement.getEndOffsetInclusive());
  }

  // --- APPEND TEXT WITH FORMATTING ---
  let cleanText = content.replace(/\*\*/g, "").replace(/^#+\s/gm, "").replace(/^\s*[\*\-]\s+/gm, "");
  const paragraphs = cleanText.split(/\n+/);

  for (let i = 0; i < paragraphs.length; i++) {
    let pText = paragraphs[i].trim();
    if (pText.length > 0) {
        let isHeading = false;

        // Check if heading
        if (pText.includes("[ГАРЧИГ]")) {
            isHeading = true;
            pText = pText.replace("[ГАРЧИГ]", "").trim();
        } else {
            // Ensure exactly one leading emoji and no internal emojis as a final safeguard
            const firstCharMatch = pText.match(/^[\u{1F300}-\u{1F6FF}\u{1F900}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F1E6}-\u{1F1FF}\u{1F200}-\u{1F2FF}]/u);
            let firstEmoji = firstCharMatch ? firstCharMatch[0] : "";
            let noEmojiText = pText.replace(/[\u{1F300}-\u{1F6FF}\u{1F900}-\u{1F9FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F1E6}-\u{1F1FF}\u{1F200}-\u{1F2FF}]/gu, "");
            pText = (firstEmoji + " " + noEmojiText).trim();
        }

        let p = body.appendParagraph(pText);
        let textObj = p.editAsText();
        applyPremiumParagraphFormatting(p, textObj, pText, isHeading, inheritAttrs);
    }
  }

  doc.saveAndClose();

  const pdfBlob = copyFile.getAs('application/pdf');
  const pdfFile = targetFolder.createFile(pdfBlob);
  pdfFile.setName(`${name} - ${PRODUCT_CONFIG.PRODUCT_NAME}.pdf`);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  copyFile.setTrashed(true);

  return {
    url: pdfFile.getUrl(),
    fileId: pdfFile.getId(),
    reused: false
  };
}

function trashDriveFileQuietly(fileId) {
  if (!fileId) return;
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (cleanupErr) {
    console.warn(`PDF cleanup failed for ${fileId}: ${cleanupErr}`);
  }
}

function sendUChatProven(userNs, pdfUrl, name, token) {
  if (!token) throw new Error("DELIVERY: UChat token байхгүй.");
  if (!userNs) throw new Error("DELIVERY: user_ns хоосон.");
  if (!pdfUrl) throw new Error("DELIVERY: PDF URL хоосон.");

  const msg = PRODUCT_CONFIG.UCHAT.DELIVERY_MESSAGE.replace("{{NAME}}", name);
  const btn = PRODUCT_CONFIG.UCHAT.DELIVERY_BTN_TEXT;

  const payload = {
    user_ns: String(userNs).trim(),
    data: {
      version: "v1",
      content: {
        messages: [
          {
            type: "text",
            text: msg,
            buttons: [
              {
                type: "url",
                caption: btn,
                url: pdfUrl
              }
            ]
          }
        ]
      }
    }
  };

  const res = UrlFetchApp.fetch(PRODUCT_CONFIG.UCHAT.ENDPOINT, {
    method: "post",
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = res.getResponseCode();
  const body = res.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error("DELIVERY HTTP " + status + ": " + body.substring(0, 200));
  }

  if (body.trim().startsWith("<")) {
    throw new Error("DELIVERY HTML response: " + body.substring(0, 120));
  }

  const json = JSON.parse(body);
  if (json.status !== "ok" && json.success !== true) {
    throw new Error("DELIVERY API failed: " + JSON.stringify(json));
  }
}

function sendErrorEmail(name, errorMsg) {
  if (PRODUCT_CONFIG.SEND_ERROR_EMAILS) {
    const adminEmail = getProperty("ADMIN_EMAIL");
    if(adminEmail) MailApp.sendEmail(adminEmail, `Error: ${PRODUCT_CONFIG.PRODUCT_NAME}`, `User: ${name}\nError: ${errorMsg}`);
  }
}

// --- КОДЫН ТӨГСГӨЛ ---
