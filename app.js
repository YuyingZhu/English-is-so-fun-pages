const STORAGE_KEY = "readingBridge.v1";
const MATERIAL_DB_NAME = "readingBridgeMaterials.v1";
const MATERIAL_DB_VERSION = 2;
const MATERIAL_STORE = "materials";
const SIGHT_WORD_STORE = "sightWordBank";
const DEV_RESET_PRACTICE_ON_LOAD = false;
const CURRENT_WEEK_ID = "2026-cw23-module-10-week-2";
const YOUNG_FRANK_CW22_TAG_MIGRATION = "young-frank-cw22-tag-fix-v1";
const CW25_MANUAL_PLAN_FIX_VERSION = "cw25-manual-weekly-plan-v3-we-are-the-future-poems";
const CW25_STORY_REWARD_AUDIO_FIX_START = "2026-06-17";
const CW26_STORY_REWARD_AUDIO_FIX_VERSION = "cw26-online-stories-v1";
const CW26_VOCABULARY_FIX_VERSION = "cw26-vocabulary-carryover-v1";
const RAZ_SUMMER_START_DATE = "2026-06-21";
const RAZ_SUMMER_END_DATE = "2026-08-30";
const RAZ_PUBLIC_LIBRARY_ENABLED = true;
const REVIEW_INTERVALS = [1, 3, 7, 14, 30];
const DEFAULT_PARENT_PASSWORD_HASH = "82e3edf5f5f3a46b5f94579b61817fd9a1f356adcef5ee22da3b96ef775c4860";
const MAX_MATERIAL_SIZE = 25 * 1024 * 1024;
const MIN_DAILY_SPELLING_WORDS = 5;
const MAX_DAILY_SPELLING_WORDS = 12;
const DAILY_SPELLING_SOURCE_RATIO = 2 / 3;
const MIN_DAILY_SPELLING_NEW_WORD_RATIO = 0.8;
const SPELLING_SELECTION_VERSION = "spelling-priority-cap-v5";
const MIN_DAILY_SIGHT_WORDS = 5;
const MAX_DAILY_SIGHT_WORDS = 10;
const WEEKLY_NEW_WORD_EXPOSURE_GOAL = 2;
const MAX_SPELLING_RETRY_CHANCES = 1;
const MAX_SPELLING_ATTEMPTS = MAX_SPELLING_RETRY_CHANCES + 1;
const STORY_REWARD_MAX_AUDIO_SECONDS = 300;
const PDFJS_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs";
const PDFJS_WORKER_URL = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs";
const XLSX_URL = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
const TESSERACT_URL = "https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js";
const WEEKLY_PLAN_OCR_VERSION = "weekly-plan-parser-v4";
const APP_ASSET_VERSION = new URL(document.currentScript?.src || window.location.href, window.location.href).searchParams.get("v") || "local";
const ALLOWED_MATERIAL_EXTENSIONS = ["jpg", "jpeg", "png", "webp", "pdf", "xlsx", "xls", "csv"];
const ALLOWED_MATERIAL_MIME_TYPES = [
  "image/jpeg",
  "image/png",
  "image/webp",
  "application/pdf",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "text/csv"
];

const learners = [
  { id: "learner-a", name: "Peggy", target: "G2 readiness" },
  { id: "learner-b", name: "Gina", target: "G2 readiness" },
  { id: "tester", name: "Tester", target: "Testing sandbox" }
];

const profiles = [
  { id: "learner-a", role: "child", defaultName: "Peggy", label: "Daily practice", avatar: "P", avatarImage: "assets/peggy-avatar.png", password: "peggy" },
  { id: "learner-b", role: "child", defaultName: "Gina", label: "Daily practice", avatar: "G", avatarImage: "assets/gina-avatar.png", password: "gina" },
  { id: "tester", role: "child", defaultName: "Tester", label: "Testing sandbox", avatar: "T", password: "" },
  { id: "parent", role: "parent", defaultName: "Parent", label: "Progress review", avatar: "P" }
];

const defaultWeeklyPlan = {
  weekId: CURRENT_WEEK_ID,
  calendarYear: 2026,
  calendarWeek: 23,
  calendarWeekLabel: "Calendar Week 23",
  archiveLabel: "2026 Calendar Week 23",
  schoolWeekLabel: "Module 10 Week 2",
  dateRange: "June 1-5, 2026",
  readingTitle: "Sky Color",
  readingArticleTitle: "Sky Color",
  readingArticleSummary: "Read realistic fiction about art, color, and creative choices. Practice setting, inferences, theme, mental images, and poetry elements.",
  phonicsFocus: "vowel pattern oo, consonant + le, blend and segment phonemes",
  grammarFocus: "informational text, biographical essay, adjectives that compare",
  uoiTopic: "tools and machines: lever, pulley, wheel, ramp, carry, easier, safe",
  masteryWords: [
    "book",
    "boyhood",
    "brook",
    "foot",
    "good",
    "hook",
    "shook",
    "took",
    "wood",
    "woo",
    "football",
    "lookout"
  ],
  sightWords: ["been", "heard", "hurry", "learn", "loved", "often", "study", "world", "good", "keep", "look", "my", "night", "took"],
  grammarPracticeWords: ["isn't", "that's", "we'll", "you'll"],
  readingSkills: ["setting", "make inferences", "theme", "create mental images", "elements of poetry"],
  weeklyVocabulary: [
    { word: "conducting", definition: "doing or leading an activity or test", example: "The class is conducting a science test." },
    { word: "lab", definition: "a room where people do science work", example: "We wore goggles in the lab." },
    { word: "method", definition: "a way to do something", example: "Her method helped her paint the sky." },
    { word: "rid", definition: "to make something go away", example: "We got rid of the old paper." },
    { word: "specimens", definition: "things used for looking at or studying", example: "The specimens were on the lab table." },
    { word: "spoiled", definition: "ruined or no longer good", example: "The rain spoiled the paper." },
    { word: "artist", definition: "a person who makes art", example: "The artist painted a mural." },
    { word: "gallery", definition: "a place where people look at art", example: "We saw paintings in the gallery." },
    { word: "merrily", definition: "in a happy way", example: "She sang merrily while she painted." },
    { word: "mural", definition: "a large picture painted on a wall", example: "The mural showed a bright sky." },
    { word: "promise", definition: "to say you will do something", example: "I promise to clean my brushes." },
    { word: "rummaged", definition: "looked through things to find something", example: "He rummaged in the box for blue paint." },
    { word: "lever", definition: "a simple machine that helps lift or move things", example: "A lever can lift a heavy lid." },
    { word: "pulley", definition: "a wheel with a rope that helps lift things", example: "The pulley lifts the bucket." },
    { word: "wheel", definition: "a round part that turns and helps things move", example: "A wheel turns and rolls." },
    { word: "ramp", definition: "a slanted path that helps things go up or down", example: "The cart rolled down the ramp." },
    { word: "carry", definition: "to hold and move something", example: "I can carry the small box." },
    { word: "easier", definition: "not as hard to do", example: "A ramp makes moving the box easier." },
    { word: "safe", definition: "not likely to hurt you", example: "A safe tool is used carefully." },
    { word: "book", definition: "pages with words or pictures to read", example: "I read a book about art." },
    { word: "look", definition: "to use your eyes to see", example: "Look at the colors in the sky." },
    { word: "took", definition: "picked up or carried away", example: "She took the paint to class." }
  ],
  epicIdeas: ["art for kids", "colors", "artists", "simple machines", "tools", "veterinarians", "animal helpers"]
};

const builtInWeeklyPlanSeeds = {
  "2026-cw25": {
    weekId: "2026-cw25",
    calendarYear: 2026,
    calendarWeek: 25,
    calendarWeekLabel: "Calendar Week 25",
    archiveLabel: "2026 Calendar Week 25",
    schoolWeekLabel: "CW25",
    dateRange: "June 15-19, 2026",
    readingTitle: "We Are the Future",
    readingArticleTitle: "We Are the Future",
    readingArticleGenre: "poetry",
    readingArticleSummary: "Read and respond to We Are the Future, a set of five HMH poetry selections. Practice noticing repeated words, rhythm, rhyme, imagery, feelings, and big ideas while using this week's vocabulary.",
    readingArticleContentStatus: "parent-provided",
    readingArticleEvidence: [
      "Me x 2 shows a speaker who reads, writes, thinks, feels, sings, and tries with a repeated times-two idea.",
      "Flying-Man imagines a flying person moving through the sky and over the sea while the speaker asks where he is going.",
      "Shoe Rack uses shoes as a picture of promise, dreams, choices, and the steps a person takes.",
      "Row, Row, Row Your Boat is a nursery rhyme about rowing gently and noticing that life can feel like a dream.",
      "Star Light, Star Bright is a nursery rhyme about seeing a bright star and making a wish."
    ].join("\n"),
    phonicsFocus: "blend phonemes, manipulate phonemes, segment phonemes, diphthongs ow/ou and oy/oi",
    grammarFocus: "opinion letter, grammar review",
    uoiTopic: "simple machines, engineering, solving problems, design, reflection, celebration",
    masteryWords: ["cow", "found", "right", "gown", "house", "how", "now", "ouch", "out", "owl", "town", "blew", "boot", "new", "too", "enjoy", "voice", "engineer", "solve", "design", "reflect", "celebrate", "favorite", "answer", "point", "walk", "watch", "where", "write", "boy", "down"],
    sightWords: ["boy", "down", "found", "how", "now", "out"],
    highFrequencyWords: ["boy", "down", "found", "how", "now", "out"],
    grammarPracticeWords: ["opinion", "reason", "because", "first", "next", "then"],
    readingSkills: ["poetry response", "repetition", "rhythm", "rhyme", "imagery", "big idea", "syllable counting", "phoneme blends"],
    uiVocabulary: [
      { word: "engineer", translation: "工程师", definition: "a person who designs or builds things to solve problems", example: "An engineer can design a useful tool." },
      { word: "solve", translation: "解决", definition: "to find an answer to a problem", example: "We can solve the problem with a simple machine." },
      { word: "design", translation: "设计", definition: "to plan or make something for a purpose", example: "We design a ramp to help the cart move." },
      { word: "reflect", translation: "反思", definition: "to think carefully about what happened or what you learned", example: "I reflect on how my tool worked." },
      { word: "celebrate", translation: "庆祝", definition: "to show happiness about something good", example: "We celebrate our hard work." },
      { word: "favorite", translation: "最喜欢的", definition: "liked best", example: "My favorite simple machine is a wheel." }
    ],
    englishVocabulary: [
      { word: "answer", translation: "答案", definition: "a reply or solution to a question", example: "I can answer the question with a detail." },
      { word: "point", translation: "指向/要点", definition: "to show where something is or an important idea", example: "Point to the detail that helps your answer." },
      { word: "right", translation: "正确的/右边", definition: "correct, or the side opposite left", example: "Choose the right answer." },
      { word: "voice", translation: "声音", definition: "the sound you use to speak", example: "Use a clear voice when you retell." },
      { word: "walk", translation: "走路", definition: "to move on foot", example: "We walk safely down the hall." },
      { word: "watch", translation: "观看/手表", definition: "to look carefully at something", example: "Watch how the wheel turns." },
      { word: "where", translation: "哪里", definition: "in or at what place", example: "Tell where the story happens." },
      { word: "write", translation: "写", definition: "to make words with letters", example: "Write one opinion sentence." },
      { word: "boy", translation: "男孩", definition: "a male child", example: "The boy found a better way." },
      { word: "down", translation: "向下", definition: "toward a lower place", example: "The truck rolls down the ramp." },
      { word: "found", translation: "发现", definition: "discovered or got something after looking", example: "She found a safe solution." },
      { word: "how", translation: "如何/怎样", definition: "in what way", example: "Explain how the machine helps." },
      { word: "now", translation: "现在", definition: "at this time", example: "Now we can test the design." },
      { word: "out", translation: "在外面/出去", definition: "away from the inside", example: "The class walked out to see the project." }
    ],
    spelling: {
      skill: "Diphthongs ow/ou and review words",
      basic: ["cow", "found", "right", "gown", "house", "how", "now", "ouch", "out", "owl", "town"],
      review: ["blew", "boot", "new", "too"],
      challenge: ["enjoy", "voice"]
    },
    weeklyVocabulary: [
      { word: "cow", definition: "a farm animal that can give milk", example: "The cow walked out of the barn." },
      { word: "found", definition: "discovered or got something after looking", example: "The boy found a safe way to move the box." },
      { word: "right", definition: "correct, or the side opposite left", example: "The right answer uses a story detail." },
      { word: "gown", definition: "a long dress or robe", example: "The gown was blue." },
      { word: "house", definition: "a building where people live", example: "The owl sat near the house." },
      { word: "how", definition: "in what way", example: "Tell how the ramp helps." },
      { word: "now", definition: "at this time", example: "Now the design works." },
      { word: "ouch", definition: "a word people say when something hurts", example: "Ouch, the box bumped my foot." },
      { word: "out", definition: "away from the inside", example: "We walked out to test the ramp." },
      { word: "owl", definition: "a bird that is often awake at night", example: "The owl looked down from the tree." },
      { word: "town", definition: "a place with streets, homes, and people", example: "The parade moved through town." },
      { word: "blew", definition: "moved air with wind or breath", example: "The wind blew the paper down." },
      { word: "boot", definition: "a strong shoe", example: "Mud was on the boot." },
      { word: "new", definition: "not old or used before", example: "We tried a new design." },
      { word: "too", definition: "also, or more than needed", example: "The ramp was too steep." },
      { word: "enjoy", definition: "to like something and feel happy about it", example: "We enjoy reading together." },
      { word: "voice", definition: "the sound you use to speak", example: "Use your voice to share an opinion." },
      { word: "engineer", definition: "a person who designs or builds things to solve problems", example: "The engineer made a useful tool." },
      { word: "solve", definition: "to find an answer to a problem", example: "We solve the problem together." },
      { word: "design", definition: "to plan or make something for a purpose", example: "The class will design a simple machine." },
      { word: "reflect", definition: "to think carefully about learning", example: "I reflect on my work." },
      { word: "celebrate", definition: "to show happiness about something good", example: "We celebrate our learning." },
      { word: "favorite", definition: "liked best", example: "This is my favorite machine." },
      { word: "answer", definition: "a reply or solution to a question", example: "My answer uses a detail." },
      { word: "point", definition: "to show where something is or an important idea", example: "Point to the evidence." },
      { word: "walk", definition: "to move on foot", example: "Walk safely in the hall." },
      { word: "watch", definition: "to look carefully at something", example: "Watch the wheel turn." },
      { word: "where", definition: "in or at what place", example: "Where does the story happen?" },
      { word: "write", definition: "to make words with letters", example: "Write your opinion." },
      { word: "boy", definition: "a male child", example: "The boy shared his idea." },
      { word: "down", definition: "toward a lower place", example: "The truck rolled down." }
    ],
    readingSelections: ["Me x 2", "Flying-Man", "Shoe Rack", "Row, Row, Row Your Boat", "Star Light, Star Bright"],
    poemSelections: [
      {
        title: "Me x 2",
        author: "Jane Medina",
        status: "parent-provided",
        evidence: "The speaker says many self-expression actions happen with a repeated times-two idea, including reading, writing, thinking, feelings, singing, trying, and doing twice as much.",
        focus: "repetition, speaker, self-expression, big idea"
      },
      {
        title: "Flying-Man",
        author: "Nursery Rhyme",
        status: "parent-provided",
        evidence: "The poem imagines a flying person in the sky and over the sea. The speaker asks where he is going and whether he can take the speaker along.",
        focus: "imagery, questions, movement, imagination"
      },
      {
        title: "Shoe Rack",
        author: "Nikki Grimes",
        status: "parent-provided",
        evidence: "A shoe rack is stacked with promise and dreams waiting to wake. The poem connects what you do, where you go, who you grow up to be, and the steps you take.",
        focus: "metaphor, choices, dreams, big idea"
      },
      {
        title: "Row, Row, Row Your Boat",
        author: "Nursery Rhyme",
        status: "public-nursery-rhyme",
        evidence: "The rhyme repeats rowing a boat gently and connects the action to a dreamlike feeling.",
        focus: "rhyme, rhythm, repetition, gentle action"
      },
      {
        title: "Star Light, Star Bright",
        author: "Nursery Rhyme",
        status: "public-nursery-rhyme",
        evidence: "The rhyme notices a bright star at night and connects it to making a wish.",
        focus: "rhyme, repetition, night imagery, wish"
      }
    ],
    jumpRopeReader: "Picture Day",
    weeklyLearningOutcomes: [
      "Review all six simple machines: wheel, pulley, lever, screw, wedge, and ramp, and explain how they make work easier.",
      "Learn to think like engineers by asking: What is the problem? What tool can help? How can we use it safely?",
      "Choose a problem and a machine to solve it, then explain how the machine helps and how to use it safely.",
      "Reflect on what was learned, share a favorite machine, and understand how humans have changed the world using tools.",
      "Read and discuss five poems: Me x 2, Flying-Man, Shoe Rack, Row Row Row Your Boat, and Star Light Star Bright.",
      "Practice poetry response by noticing repeated words, rhythm, rhyme, imagery, feelings, and big ideas.",
      "Learn words with diphthongs ow/ou, including basic, review, and challenge spelling words.",
      "Review syllable counting and phoneme blends.",
      "Read Jump Rope Reader Picture Day.",
      "Practice opinion letter writing and grammar review.",
      "Complete Friday dictation practice."
    ],
    homeLearningSuggestions: [
      "Ask your child to walk around the house or look in a book or magazine and find objects or words with the /ow/ sound.",
      "Create two columns on paper, ow like cow and ou like house, then sort basic, review, and challenge words."
    ],
    epicIdeas: ["poems", "Me x 2", "Flying-Man", "Shoe Rack", "Row Row Row Your Boat", "Star Light Star Bright", "Picture Day", "simple machines", "engineering", "opinion writing", "diphthongs ow/ou"],
    sourceTitle: "Built-in Calendar Week 25 recovery plan",
    updatedAt: "2026-06-14T00:00:00.000Z",
    listeningStories: cw25ListeningStories()
  },
  "2026-cw24": {
    weekId: "2026-cw24",
    calendarYear: 2026,
    calendarWeek: 24,
    calendarWeekLabel: "Calendar Week 24",
    archiveLabel: "2026 Calendar Week 24",
    schoolWeekLabel: "Module 10 Week 3",
    dateRange: "June 8-12, 2026",
    readingTitle: "Joaquin's Zoo",
    readingArticleTitle: "Joaquin's Zoo",
    readingArticleSummary: "Read fantasy fiction about Joaquin imagining animals and objects in a zoo. Practice retelling, problem and solution, and fantasy details.",
    phonicsFocus: "long oo sound and story words",
    grammarFocus: "adverbs and complete sentence answers",
    uoiTopic: "animals, imagination, and describing objects",
    masteryWords: ["moon", "soon", "zoo", "object", "problem", "neat", "dream", "blew"],
    sightWords: ["because", "before", "after", "first", "next", "then"],
    grammarPracticeWords: ["usually", "quickly", "slowly", "carefully"],
    readingSkills: ["retell", "problem and solution", "fantasy details", "make inferences"],
    uiVocabulary: [
      { word: "tool", translation: "工具", definition: "something people use to do work", example: "A tool can help us solve a problem." },
      { word: "object", translation: "物体", definition: "a thing that you can see or touch", example: "The object has a handle." },
      { word: "scissors", translation: "剪刀", definition: "a tool used for cutting", example: "Scissors can cut paper." },
      { word: "screw", translation: "螺丝", definition: "a small metal object that holds things together", example: "A screw can hold the handle in place." },
      { word: "handle", translation: "手柄", definition: "the part you hold with your hand", example: "I hold the handle to open the door." },
      { word: "problem", translation: "问题", definition: "something that needs to be solved", example: "The machine helps solve a problem." }
    ],
    englishVocabulary: [
      { word: "attitude", translation: "态度", definition: "the way someone thinks or feels about something", example: "A good attitude helps you keep trying." },
      { word: "bounds", translation: "界限", definition: "limits or edges", example: "The animal stayed inside the bounds." },
      { word: "cab", translation: "出租车", definition: "a taxi", example: "They rode in a cab." },
      { word: "dream", translation: "梦想", definition: "something you hope for or imagine", example: "Joaquin had a dream about a zoo." },
      { word: "instant", translation: "瞬间", definition: "a very short moment", example: "In an instant, the room looked different." },
      { word: "instructor", translation: "指导者", definition: "a person who teaches or guides", example: "The instructor helped the class." },
      { word: "build", translation: "建造", definition: "to make something", example: "We can build a small zoo model." },
      { word: "golden", translation: "金色", definition: "bright yellow like gold", example: "The golden light shone on the zoo." },
      { word: "neat", translation: "整齐", definition: "clean and in order", example: "Joaquin made a neat drawing." },
      { word: "scraps", translation: "碎片", definition: "small leftover pieces", example: "He used scraps of paper." },
      { word: "usually", translation: "经常", definition: "most of the time", example: "Joaquin usually draws animals." }
    ],
    spelling: {
      skill: "Vowel Patterns /oo/",
      basic: ["blew", "boot", "moon", "new", "noon", "soon", "soup", "too", "you", "zoo"],
      review: ["book", "boyhood", "brook", "foot"],
      challenge: ["balloon", "shampoo"]
    },
    highFrequencyWords: ["bear", "color", "happy", "money", "music", "second", "sound", "without"],
    weeklyVocabulary: [
      { word: "zoo", definition: "a place where people can see animals", example: "Joaquin imagined a zoo." },
      { word: "object", definition: "a thing that you can see or touch", example: "The object looked like an animal." },
      { word: "problem", definition: "something that needs to be solved", example: "The story problem made Joaquin think." },
      { word: "neat", definition: "clean, tidy, or cool", example: "The neat drawing looked like a zoo animal." },
      { word: "dream", definition: "to imagine something while sleeping or thinking", example: "I dream about a moonlit zoo." },
      { word: "moon", definition: "the bright round object we see in the night sky", example: "The moon shone over the zoo." },
      { word: "soon", definition: "after a short time", example: "Soon the story will begin." },
      { word: "blew", definition: "moved air with wind or breath", example: "The wind blew past the animals." },
      { word: "usually", definition: "most of the time", example: "Joaquin usually draws animals." },
      { word: "quickly", definition: "in a fast way", example: "The animal moved quickly." },
      { word: "slowly", definition: "in a slow way", example: "The turtle walked slowly." },
      { word: "carefully", definition: "in a careful way", example: "Joaquin looked carefully at the object." },
      { word: "because", definition: "for the reason that", example: "Joaquin smiled because the object looked like an animal." },
      { word: "before", definition: "earlier than something", example: "Before the story ended, Joaquin had a new idea." },
      { word: "after", definition: "later than something", example: "After he looked again, he saw a zoo." },
      { word: "first", definition: "coming before the others", example: "First, Joaquin looked at the object." },
      { word: "next", definition: "coming right after", example: "Next, he imagined an animal." },
      { word: "then", definition: "after that", example: "Then the room felt like a zoo." },
      { word: "animals", definition: "living things that are not plants", example: "The zoo had many animals." },
      { word: "imagination", definition: "the power to make pictures or ideas in your mind", example: "Joaquin used imagination to make a zoo." },
      { word: "describing", definition: "telling what something is like", example: "We are describing the object." }
    ],
    epicIdeas: ["zoo animals", "fantasy stories", "animal imagination", "moon stories", "problem and solution"],
    sourceTitle: "Built-in Calendar Week 24 plan",
    weeklySequence: {
      sourceType: "weekly sequence",
      sequenceLabel: "Module 10 Week 3",
      selections: ["Joaquin's Zoo (Fantasy)"],
      vocabulary: {
        oralPower: ["attitude", "bounds", "cab", "dream", "instant", "instructor"],
        power: ["build", "golden", "neat", "scraps", "usually"]
      },
      generativeVocabulary: ["prefix re-"],
      readingWorkshop: ["Text Organization", "Make Connections", "Characters", "Central Idea"],
      foundationalSkills: ["Blend Phonemes", "Segment Phonemes", "Vowel Pattern /oo/"],
      highFrequencyWords: ["bear", "color", "happy", "money", "music", "second", "sound", "without"],
      decodableWords: ["blue", "new", "soon", "too", "try", "you"],
      spelling: {
        skill: "Vowel Patterns /oo/",
        basic: ["blew", "boot", "moon", "new", "noon", "soon", "soup", "too", "you", "zoo"],
        review: ["book", "boyhood", "brook", "foot"],
        challenge: ["balloon", "shampoo"]
      },
      jumpRopeReader: "The Monsoon",
      writingWorkshopGrammar: ["Informational Text: Biographical Essay", "Spelling"]
    },
    listeningStories: []
  }
};

const knownUoiVocabulary = [
  "push",
  "pull",
  "lift",
  "easier",
  "ramp",
  "lever",
  "pulley",
  "screw",
  "wedge",
  "wheel",
  "carry",
  "safe",
  "floor",
  "model",
  "real",
  "straight",
  "whole",
  "designed",
  "design",
  "engineer",
  "solve",
  "reflect",
  "celebrate",
  "favorite",
  "favourite",
  "dancing",
  "fee",
  "float",
  "training",
  "whatever"
];

const phonicsExampleWords = [
  "catch",
  "catches",
  "catcher",
  "catching",
  "caught",
  "cents",
  "long",
  "thank",
  "house",
  "ouch"
];

const skillPath = [
  {
    id: "g1",
    title: "G1 Foundation",
    subtitle: "American early literacy",
    skills: [
      "Blend and segment sounds",
      "Read short and long vowel words",
      "Decode CVC and final-e words",
      "Read vowel teams and r-controlled words",
      "Count syllables and decode two-syllable words",
      "Read sight words with automaticity",
      "Read decodable text with expression",
      "Write complete sentences"
    ]
  },
  {
    id: "g2",
    title: "G2 Readiness",
    subtitle: "Entrance exam application",
    skills: [
      "Decode open, closed, and consonant-le syllables",
      "Use prefixes, suffixes, and inflectional endings",
      "Read words with variable vowel teams",
      "Use context clues for new words",
      "Answer main idea, detail, sequence, and cause/effect questions",
      "Use pronouns, possessives, verb tense, adverbs, and punctuation",
      "Write short answers in complete sentences"
    ]
  },
  {
    id: "g3",
    title: "Future G3 Prep",
    subtitle: "Reserved for later papers",
    skills: [
      "Decode multisyllabic words",
      "Study roots, prefixes, and suffixes",
      "Read longer fiction and nonfiction",
      "Explain evidence from text",
      "Write organized multi-sentence responses"
    ]
  }
];

const wordHelp = {
  designed: ["made a plan for how something should look or work", "She designed a small house with blocks."],
  floor: ["the flat part of a room that you walk on", "The toy car rolled across the floor."],
  model: ["a small copy of something bigger", "He made a model of a bridge."],
  real: ["not pretend", "A real bird can fly."],
  straight: ["not bent or curved", "Draw a straight line from the door to the tree."],
  whole: ["all of something", "She read the whole book."],
  dancing: ["moving your body to music", "The children are dancing after class."],
  fee: ["money paid for something", "The club has a small fee."],
  float: ["to stay on top of water or air", "The leaf can float on the pond."],
  training: ["practice that helps you get better", "Soccer training made him faster."],
  whatever: ["anything or no matter what", "You can choose whatever book you like."],
  "can't": ["a short way to say cannot", "I can't find my pencil."],
  "didn't": ["a short way to say did not", "She didn't see the dog."],
  "i'll": ["a short way to say I will", "I'll read one more page."],
  "i'm": ["a short way to say I am", "I'm ready to spell."],
  "isn't": ["a short way to say is not", "The box isn't heavy."],
  "it's": ["a short way to say it is", "It's time to read."],
  "that's": ["a short way to say that is", "That's my favorite word."],
  "wasn't": ["a short way to say was not", "The game wasn't hard."],
  "we'll": ["a short way to say we will", "We'll practice again tomorrow."],
  "you'll": ["a short way to say you will", "You'll know this word soon."],
  push: ["to move something away from you", "Push the door to open it."],
  pull: ["to move something toward you", "Pull the wagon up the hill."],
  lift: ["to move something up", "Lift the box with two hands."],
  easier: ["not as hard to do", "A ramp makes moving the box easier."],
  ramp: ["a slanted path that helps things go up or down", "The car rolled down the ramp."],
  lever: ["a simple machine that helps lift or move things", "A seesaw is like a lever."],
  pulley: ["a wheel with a rope that helps lift things", "A pulley can lift a bucket."],
  wheel: ["a round part that turns and helps things move", "The wheel turns on the bike."],
  carry: ["to hold and move something", "I can carry the small box."],
  safe: ["not likely to hurt you", "A safe tool is used carefully."],
  conducting: ["doing or leading an activity or test", "The class is conducting a science test."],
  lab: ["a room where people do science work", "We wore goggles in the lab."],
  method: ["a way to do something", "Her method helped her paint the sky."],
  rid: ["to make something go away", "We got rid of the old paper."],
  specimens: ["things used for looking at or studying", "The specimens were on the lab table."],
  spoiled: ["ruined or no longer good", "The rain spoiled the paper."],
  artist: ["a person who makes art", "The artist painted a mural."],
  gallery: ["a place where people look at art", "We saw paintings in the gallery."],
  merrily: ["in a happy way", "She sang merrily while she painted."],
  mural: ["a large picture painted on a wall", "The mural showed a bright sky."],
  promise: ["to say you will do something", "I promise to clean my brushes."],
  rummaged: ["looked through things to find something", "He rummaged in the box for blue paint."],
  book: ["pages with words or pictures to read", "I read a book about art."],
  boyhood: ["the time when a person was a boy", "He loved drawing during his boyhood."],
  brook: ["a small stream of water", "The brook ran beside the trail."],
  foot: ["the part of your body you stand on", "My foot stepped in the mud."],
  good: ["nice, helpful, or right", "That is a good idea."],
  hook: ["a curved thing used to hold or catch", "Hang the bag on the hook."],
  shook: ["moved back and forth quickly", "The wet dog shook water off."],
  took: ["picked up or carried away", "She took the paint to class."],
  wood: ["hard material from a tree", "The table is made of wood."],
  woo: ["to try to win someone over", "The bird sang to woo a mate."],
  football: ["a game played with a ball by two teams", "They played football after school."],
  lookout: ["a place or person that watches for something", "The lookout saw the boat first."],
  catch: ["to get or stop something moving", "I can catch the ball."],
  catches: ["gets or stops something moving", "She catches the ball."],
  catcher: ["a person who catches something", "The catcher catches the ball."],
  catching: ["getting or stopping something that is moving", "She is catching the ball."],
  caught: ["got or stopped something that was moving", "He caught the ball."],
  cents: ["small coins or parts of a dollar", "I have ten cents."],
  long: ["going far from end to end", "The rope is long."],
  thank: ["to tell someone you are glad for help", "I thank my brother."],
  house: ["a home where people live", "The house has a red door."],
  ouch: ["a word people say when something hurts", "Ouch, my finger hurts."],
  syllable: ["a word part with one vowel sound", "Rabbit has two syllables: rab-bit."],
  decode: ["to read a word by using sounds and word parts", "I can decode sunshine by reading sun and shine."],
  epic: ["very big or exciting", "The story felt epic because the trip was so long."]
};

const foundationalLearningTerms = new Set([
  "syllable",
  "syllables",
  "decode",
  "decoding",
  "contraction",
  "contractions",
  "prefix",
  "prefixes",
  "suffix",
  "suffixes",
  "blend",
  "blends",
  "segment",
  "segments",
  "phonics",
  "foundational"
]);

const sightWordSeedBank = [
  ...[
    "the", "of", "and", "a", "to", "in", "is", "you", "that", "it", "he", "was", "for", "on", "are", "as",
    "with", "his", "they", "I", "at", "be", "this", "have", "from", "or", "one", "had", "by", "words", "but",
    "not", "what", "all", "were", "we", "when", "your", "can", "said", "there", "use", "an", "each", "which",
    "she", "do", "how", "their", "if", "will", "up", "other", "about", "out", "many", "then", "them", "these",
    "so", "some", "her", "would", "make", "like", "him", "into", "time", "has", "look", "two", "more", "write",
    "go", "see", "number", "no", "way", "could", "people", "my", "than", "first", "water", "been", "call",
    "who", "oil", "its", "now", "find", "long", "down", "day", "did", "get", "come", "made", "may", "part"
  ].map((word, index) => ({
    word,
    grade: index < 60 ? "G1" : "G2",
    source: "Fry first 100",
    isIrregular: ["the", "of", "to", "said", "was", "were", "have", "one", "two", "would", "could", "people"].includes(word),
    heartPart: ["the", "of", "said", "one", "two", "would", "could", "people"].includes(word) ? "tricky sound" : "",
    definition: wordHelp[normalizeWord(word)]?.[0] ?? "a word good readers know quickly",
    example: wordHelp[normalizeWord(word)]?.[1] ?? `I can read the word ${word}.`
  })),
  ...[
    "over", "new", "sound", "take", "only", "little", "work", "know", "place", "year", "live", "me",
    "back", "give", "most", "very", "after", "thing", "our", "just", "name", "good", "sentence", "man",
    "think", "say", "great", "where", "help", "through", "much", "before", "line", "right", "too", "mean",
    "old", "any", "same", "tell", "boy", "follow", "came", "want", "show", "also", "around", "form",
    "three", "small", "set", "put", "end", "does", "another", "well", "large", "must", "big", "even",
    "such", "because", "turn", "here", "why", "ask", "went", "men", "read", "need", "land", "different",
    "home", "us", "move", "try", "kind", "hand", "again", "change", "off", "play", "spell", "air",
    "away", "animal", "point", "page", "letter", "mother", "answer", "found", "study", "still", "learn",
    "should", "world"
  ].map((word) => ({
    word,
    grade: "G2",
    source: "Fry second 100",
    isIrregular: ["only", "know", "give", "very", "through", "does", "because", "where", "again", "should"].includes(word),
    heartPart: ["know", "through", "does", "should"].includes(word) ? "tricky sound" : "",
    definition: wordHelp[normalizeWord(word)]?.[0] ?? "a word good readers know quickly",
    example: wordHelp[normalizeWord(word)]?.[1] ?? `I can read the word ${word}.`
  })),
  ...[
    "begin", "brother", "front", "picture", "room", "someone", "sometimes", "young",
    "can't", "didn't", "I'll", "I'm", "isn't", "it's", "that's", "wasn't", "we'll", "you'll"
  ].map((word) => ({
    word,
    grade: "G1-G2",
    source: "HMH weekly high frequency / contractions",
    isIrregular: true,
    heartPart: word.includes("'") ? "apostrophe contraction" : "",
    definition: wordHelp[normalizeWord(word)]?.[0] ?? "a weekly word to read quickly",
    example: wordHelp[normalizeWord(word)]?.[1] ?? `Read ${word} in a sentence.`
  }))
];

sightWordSeedBank.forEach((item) => {
  const key = normalizeWord(item.word);
  if (key && !wordHelp[key]) wordHelp[key] = [item.definition, item.example];
});

const spellingExampleSentences = {
  balloon: "The red balloon floated above the party.",
  blew: "The wind blew the paper across the yard.",
  book: "She opened the book and read the story.",
  boot: "Mud stuck to the bottom of my boot.",
  boyhood: "He played near the brook during his boyhood.",
  brook: "A small brook ran beside the path.",
  build: "We can build a tower with blocks.",
  cab: "The cab stopped at the corner.",
  cow: "The cow ate grass in the field.",
  dream: "She had a dream about a golden zoo.",
  enjoy: "We enjoy reading poems together.",
  foot: "My foot left a print in the sand.",
  found: "I found my pencil under the desk.",
  golden: "The golden sun shone on the wall.",
  gown: "The dancer wore a long blue gown.",
  handle: "Hold the handle to open the door.",
  house: "Our house has a red door.",
  how: "Tell me how you solved the problem.",
  instant: "In an instant, the room became quiet.",
  instructor: "The instructor showed us how to use the tool.",
  moon: "The moon looked bright in the night sky.",
  neat: "He kept his desk neat and clean.",
  new: "She wore her new shoes to school.",
  now: "Now it is time to read.",
  noon: "We eat lunch at noon.",
  object: "The object on the table was a small tool.",
  ouch: "Ouch, my finger hurts.",
  out: "The children went out to play.",
  owl: "The owl sat in the tree at night.",
  problem: "A ramp can solve the problem of moving a heavy box.",
  right: "That is the right answer.",
  scissors: "Use scissors to cut the paper carefully.",
  scraps: "She saved the paper scraps for an art project.",
  screw: "A screw can hold two pieces of wood together.",
  shampoo: "He washed his hair with shampoo.",
  soon: "The bus will come soon.",
  soup: "The hot soup filled the bowl.",
  too: "The box is too heavy to lift alone.",
  tool: "A hammer is a tool for building.",
  town: "Our town has a library and a park.",
  usually: "She usually reads before bedtime.",
  voice: "Use a clear voice when you read aloud.",
  you: "You can choose a book from the shelf.",
  zoo: "The zoo has animals from many places."
};

const missionCenters = [
  {
    id: "listen",
    minutes: 2,
    title: "Listen & Spell",
    domain: "Listening",
    prompt: "Hear weekly words, then spell them without looking.",
    skill: "accurate listening and word recognition",
    visual: "listen",
    practiceType: "spelling",
    visualIcon: "🎧",
    theme: "sky",
    completionGoal: 5,
    progressUnit: "words"
  },
  {
    id: "sight",
    minutes: 2,
    title: "Sight Word Center",
    domain: "Sight Words",
    prompt: "Read contractions fast, then use one in a sentence.",
    skill: "automaticity with irregular and high-frequency words",
    visual: "sight",
    practiceType: "flashcards",
    visualIcon: "★",
    theme: "coral",
    completionGoal: 8,
    progressUnit: "words"
  },
  {
    id: "wordwork",
    minutes: 3,
    title: "Grammar Practice",
    domain: "Grammar",
    prompt: "Answer short questions using this week's grammar focus.",
    skill: "grammar and complete sentence thinking",
    visual: "blocks",
    practiceType: "grammar",
    visualIcon: "ab",
    theme: "green",
    completionGoal: 3,
    progressUnit: "questions"
  },
  {
    id: "speak",
    minutes: 3,
    title: "Speaking Retell",
    domain: "Speaking",
    prompt: "Read one sentence aloud. Record it, replay it, and try for smooth expression.",
    skill: "fluency and oral language",
    visual: "speaking",
    practiceType: "speaking",
    visualIcon: "🎙",
    theme: "peach",
    completionGoal: 1,
    progressUnit: "recording"
  },
  {
    id: "reading",
    minutes: 2,
    title: "Shared Reading",
    domain: "Reading",
    prompt: "Read a short passage and answer one question.",
    skill: "comprehension and self-correction",
    visual: "reading",
    practiceType: "reading",
    visualIcon: "🔎",
    theme: "mint",
    completionGoal: 1,
    progressUnit: "answer"
  },
  {
    id: "writing",
    minutes: 2,
    title: "Writing Response",
    domain: "Writing",
    prompt: "Write one complete sentence using a weekly word.",
    skill: "sentence writing and punctuation",
    visual: "writing",
    practiceType: "writing",
    visualIcon: "✎",
    theme: "blue",
    completionGoal: 1,
    progressUnit: "sentence"
  },
  {
    id: "warmup",
    minutes: 1,
    title: "Warm-Up",
    domain: "Review",
    prompt: "Review one word or sound that needs more practice.",
    skill: "spaced review",
    visual: "warmup",
    practiceType: "review",
    visualIcon: "↻",
    theme: "yellow",
    completionGoal: 3,
    progressUnit: "reviews"
  }
];

const dailyMissionCenterIds = ["reading", "speak", "listen", "sight", "wordwork"];

const adventureLevelDefinitions = [
  { id: "read-warmup", level: 1, title: "Listen & Read" },
  { id: "speak", level: 2, title: "Retell Story" },
  { id: "word-power", level: 3, title: "Word Power" },
  { id: "listen", level: 4, title: "Spelling Challenge" },
  { id: "sight", level: 5, title: "Sight Word Sprint" },
  { id: "wordwork", level: 6, title: "Grammar Builder" },
  { id: "story-reward", level: 7, title: "Story Reward" }
];

const adventureLevelIds = adventureLevelDefinitions.map((level) => level.id);
const legacyCenterAdventureMap = {
  speak: "speak",
  listen: "listen",
  sight: "sight",
  wordwork: "wordwork"
};

const practiceDomains = [
  {
    id: "words",
    title: "Words",
    items: [
      ["Sight word speed", "Read five sight words. Tap I know it only when you can read each one quickly."],
      ["Listen and spell", "Listen to a weekly mastery word, then type it without looking."],
      ["Word meaning", "Pick a word and say a kid-friendly definition in English."]
    ]
  },
  {
    id: "phonics",
    title: "Phonics",
    items: [
      ["Long vowel check", "Sort words into short vowel, final-e, and vowel team groups."],
      ["Syllable split", "Break two-syllable words into readable parts."],
      ["Decode line", "Read a line of words, then read the same pattern in a sentence."]
    ]
  },
  {
    id: "reading",
    title: "Reading",
    items: [
      ["Main idea", "Read a short text and tell what it is mostly about."],
      ["Text evidence", "Find one detail that proves your answer."],
      ["Retell", "Say the beginning, middle, and end in your own words."]
    ]
  },
  {
    id: "grammar",
    title: "Grammar",
    items: [
      ["Contractions", "Match each contraction to its full form."],
      ["Pronouns", "Choose the pronoun that makes the sentence clear."],
      ["Punctuation", "Add end marks, commas, or quotation marks."]
    ]
  }
];

const examItems = [
  {
    title: "Phonics",
    prompt: "Which word has a long a sound?",
    choices: ["cat", "cake", "cap"],
    answer: "cake"
  },
  {
    title: "Vocabulary",
    prompt: "A ramp helps something go...",
    choices: ["up or down", "under water", "inside a bag"],
    answer: "up or down"
  },
  {
    title: "Grammar",
    prompt: "Choose the best word: She ___ a model.",
    choices: ["have", "has", "had"],
    answer: "has"
  },
  {
    title: "Comprehension",
    prompt: "If a story tells about a real person&apos;s life, it is a...",
    choices: ["biography", "fable", "fantasy"],
    answer: "biography"
  }
];

const epicSeedLibrary = [
  {
    title: "Simple Machines: Lever",
    author: "Public Epic listing",
    source: "Epic public catalog",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Nonfiction",
    readTime: "10 min",
    topicTags: ["simple machines", "lever", "science", "push", "pull"],
    skillTags: ["cause and effect", "domain vocabulary", "nonfiction details"],
    searchTerms: ["simple machines lever", "lever for kids", "push pull lever"]
  },
  {
    title: "Simple Machines",
    author: "Public Epic listing",
    source: "Epic public catalog",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Nonfiction",
    readTime: "10-15 min",
    topicTags: ["simple machines", "wheel", "ramp", "lever", "science"],
    skillTags: ["main idea", "details", "cause and effect"],
    searchTerms: ["simple machines", "ramps levers wheels", "how machines work"]
  },
  {
    title: "Push and Pull",
    author: "Early science reader",
    source: "Curated public book-list metadata",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Nonfiction",
    readTime: "8-12 min",
    topicTags: ["push", "pull", "motion", "force", "science"],
    skillTags: ["cause and effect", "vocabulary", "complete sentence answers"],
    searchTerms: ["push and pull", "force and motion kids", "push pull science"]
  },
  {
    title: "Ramps and Wedges",
    author: "Early science reader",
    source: "Curated public book-list metadata",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Nonfiction",
    readTime: "8-12 min",
    topicTags: ["ramp", "inclined plane", "simple machines", "science"],
    skillTags: ["domain vocabulary", "sequence", "nonfiction details"],
    searchTerms: ["ramps for kids", "inclined plane", "simple machines ramp"]
  },
  {
    title: "Wheels at Work",
    author: "Early science reader",
    source: "Curated public book-list metadata",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Nonfiction",
    readTime: "8-12 min",
    topicTags: ["wheel", "transportation", "simple machines", "motion"],
    skillTags: ["main idea", "details", "vocabulary"],
    searchTerms: ["wheels at work", "wheels for kids", "wheel simple machine"]
  },
  {
    title: "What Do Architects Do?",
    author: "STEM picture nonfiction",
    source: "Curated public book-list metadata",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Nonfiction",
    readTime: "10-15 min",
    topicTags: ["architecture", "buildings", "design", "model"],
    skillTags: ["main idea", "details", "domain vocabulary"],
    searchTerms: ["architects for kids", "architecture for kids", "buildings design"]
  },
  {
    title: "Young Frank, Architect",
    author: "Frank Viva",
    source: "School weekly title",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Fiction",
    readTime: "10-15 min",
    topicTags: ["architecture", "design", "buildings", "model"],
    skillTags: ["retell", "character", "story sequence"],
    searchTerms: ["Young Frank Architect", "Frank Viva architect", "architecture story kids"]
  },
  {
    title: "Iggy Peck, Architect",
    author: "Andrea Beaty",
    source: "Public children's book lists",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Fiction",
    readTime: "10-15 min",
    topicTags: ["architecture", "buildings", "creative thinking", "design"],
    skillTags: ["retell", "problem solution", "character traits"],
    searchTerms: ["Iggy Peck Architect", "architect story", "building story kids"]
  },
  {
    title: "Rosie Revere, Engineer",
    author: "Andrea Beaty",
    source: "Public children's book lists",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Fiction",
    readTime: "10-15 min",
    topicTags: ["inventors", "engineering", "design", "machines"],
    skillTags: ["problem solution", "retell", "character"],
    searchTerms: ["Rosie Revere Engineer", "engineering for kids", "inventor story"]
  },
  {
    title: "Long Vowels",
    author: "Decodable phonics reader",
    source: "Epic/public phonics listing metadata",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Decodable",
    readTime: "8-10 min",
    topicTags: ["phonics", "long vowels", "decodable"],
    skillTags: ["long vowels", "decode words", "fluency"],
    searchTerms: ["long vowels", "long vowel decodable", "phonics long vowels"]
  },
  {
    title: "Vowel Teams",
    author: "Decodable phonics reader",
    source: "Epic/public phonics listing metadata",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Decodable",
    readTime: "8-10 min",
    topicTags: ["phonics", "vowel teams", "decodable"],
    skillTags: ["vowel teams", "decode words", "word recognition"],
    searchTerms: ["vowel teams", "ai ay ee ea", "decodable vowel teams"]
  },
  {
    title: "Contractions",
    author: "Grammar early reader",
    source: "Curated public book-list metadata",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Language Arts",
    readTime: "8-10 min",
    topicTags: ["contractions", "grammar", "sight words"],
    skillTags: ["contractions", "sight words", "sentence writing"],
    searchTerms: ["contractions", "can't didn't we'll", "grammar contractions kids"]
  },
  {
    title: "Syllables",
    author: "Phonics early reader",
    source: "Curated public book-list metadata",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Language Arts",
    readTime: "8-10 min",
    topicTags: ["syllables", "phonics", "word parts"],
    skillTags: ["syllable division", "decode words", "fluency"],
    searchTerms: ["syllables", "word parts", "two syllable words"]
  },
  {
    title: "Main Idea and Details",
    author: "Reading skills reader",
    source: "Curated public book-list metadata",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Reading Skills",
    readTime: "10 min",
    topicTags: ["reading comprehension", "nonfiction", "school"],
    skillTags: ["main idea", "details", "complete sentence answers"],
    searchTerms: ["main idea details", "reading comprehension grade 1", "nonfiction main idea"]
  },
  {
    title: "Cause and Effect",
    author: "Reading skills reader",
    source: "Curated public book-list metadata",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Reading Skills",
    readTime: "10 min",
    topicTags: ["reading comprehension", "science", "how things work"],
    skillTags: ["cause and effect", "complete sentence answers"],
    searchTerms: ["cause and effect kids", "reading cause effect", "why because grade 1"]
  },
  {
    title: "Animal Homes",
    author: "Early nonfiction reader",
    source: "Public G1/G2 reading lists",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Nonfiction",
    readTime: "10-15 min",
    topicTags: ["animals", "homes", "habitats", "nonfiction"],
    skillTags: ["main idea", "details", "vocabulary"],
    searchTerms: ["animal homes", "habitats for kids", "where animals live"]
  },
  {
    title: "How Things Move",
    author: "Early science reader",
    source: "Public G1/G2 reading lists",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Nonfiction",
    readTime: "10-15 min",
    topicTags: ["motion", "push", "pull", "science"],
    skillTags: ["cause and effect", "domain vocabulary"],
    searchTerms: ["how things move", "motion for kids", "push pull move"]
  },
  {
    title: "Building a House",
    author: "Early nonfiction reader",
    source: "Public G1/G2 reading lists",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: "Nonfiction",
    readTime: "10-15 min",
    topicTags: ["buildings", "architecture", "design", "tools"],
    skillTags: ["sequence", "main idea", "vocabulary"],
    searchTerms: ["building a house", "how houses are built", "buildings for kids"]
  }
];

const epicSearchLibrary = [
  ["Decodable: Short Vowels", ["phonics", "short vowels", "decodable"], ["short vowels", "cvc words", "short vowel decodable"]],
  ["Decodable: Long A", ["phonics", "long vowels", "final e"], ["long a", "a_e words", "long vowel a"]],
  ["Decodable: Long E", ["phonics", "long vowels", "vowel teams"], ["long e", "ee ea words", "long vowel e"]],
  ["Decodable: Long I", ["phonics", "long vowels", "final e"], ["long i", "i_e words", "long vowel i"]],
  ["Decodable: Long O", ["phonics", "long vowels", "vowel teams"], ["long o", "o_e words", "oa ow words"]],
  ["Decodable: Long U", ["phonics", "long vowels", "final e"], ["long u", "u_e words", "long vowel u"]],
  ["R-Controlled Vowels", ["phonics", "r-controlled vowels", "decode"], ["ar er ir or ur", "r controlled vowels", "bossy r"]],
  ["Vowel Teams: AI and AY", ["phonics", "vowel teams", "decode"], ["ai ay words", "long a vowel teams", "vowel teams ai ay"]],
  ["Vowel Teams: EE and EA", ["phonics", "vowel teams", "decode"], ["ee ea words", "long e vowel teams", "vowel teams ee ea"]],
  ["Two-Syllable Words", ["phonics", "syllables", "decode"], ["two syllable words", "syllable division", "decode syllables"]],
  ["Compound Words", ["phonics", "word parts", "decode"], ["compound words", "word parts", "sunshine playground"]],
  ["Prefixes: un-", ["phonics", "prefix", "word meaning"], ["prefix un", "unhappy untie undo", "prefixes for kids"]],
  ["Suffixes: -ed and -ing", ["phonics", "suffix", "inflectional endings"], ["ed ing endings", "suffixes ed ing", "inflectional endings"]],
  ["Contractions Practice", ["grammar", "contractions", "sight words"], ["contractions", "can't didn't isn't", "apostrophe contractions"]],
  ["Sight Words Grade 1", ["sight words", "fluency", "word recognition"], ["grade 1 sight words", "high frequency words", "sight word readers"]],
  ["Sight Words Grade 2", ["sight words", "fluency", "word recognition"], ["grade 2 sight words", "high frequency words grade 2", "sight word stories"]],
  ["Main Idea Readers", ["reading comprehension", "main idea", "details"], ["main idea grade 1", "main idea details", "nonfiction main idea"]],
  ["Sequence Readers", ["reading comprehension", "sequence", "retell"], ["sequence story", "first next last", "retell beginning middle end"]],
  ["Cause and Effect Readers", ["reading comprehension", "cause and effect", "science"], ["cause and effect", "why because", "science cause effect"]],
  ["Problem and Solution Stories", ["reading comprehension", "problem solution", "fiction"], ["problem solution story", "fiction problem solution", "story elements"]],
  ["Character Traits Stories", ["reading comprehension", "character", "fiction"], ["character traits", "story character", "fiction grade 1"]],
  ["Nonfiction Text Features", ["reading comprehension", "nonfiction", "details"], ["nonfiction text features", "captions diagrams labels", "informational text"]],
  ["Animal Habitats", ["animals", "habitats", "nonfiction"], ["animal habitats", "where animals live", "animal homes"]],
  ["Life Cycles", ["science", "life cycles", "nonfiction"], ["life cycle", "butterfly life cycle", "frog life cycle"]],
  ["Weather and Seasons", ["science", "weather", "nonfiction"], ["weather for kids", "seasons for kids", "clouds rain sun"]],
  ["Plants Grow", ["science", "plants", "nonfiction"], ["plants grow", "seeds and plants", "plant life cycle"]],
  ["Matter and Materials", ["science", "matter", "nonfiction"], ["matter for kids", "solid liquid gas", "materials for kids"]],
  ["Magnets", ["science", "force", "nonfiction"], ["magnets for kids", "magnet science", "push pull magnets"]],
  ["Sound and Light", ["science", "sound", "light"], ["sound light for kids", "light science", "sound science"]],
  ["Maps and Places", ["social studies", "maps", "community"], ["maps for kids", "community places", "where we live"]],
  ["Community Helpers", ["social studies", "community", "nonfiction"], ["community helpers", "jobs for kids", "helpers in town"]],
  ["School Stories", ["fiction", "school", "retell"], ["school stories", "first grade stories", "classroom story"]],
  ["Friendship Stories", ["fiction", "friendship", "character"], ["friendship stories", "friends at school", "kindness story"]],
  ["Funny Early Readers", ["fiction", "fluency", "retell"], ["funny early readers", "funny grade 1 books", "easy funny books"]],
  ["STEM Inventors", ["inventors", "engineering", "nonfiction"], ["inventors for kids", "engineering for kids", "stem books kids"]],
  ["Buildings and Bridges", ["architecture", "buildings", "design"], ["buildings and bridges", "bridges for kids", "architecture kids"]],
  ["Tools and Building", ["tools", "buildings", "science"], ["tools for kids", "building things", "construction for kids"]],
  ["Simple Machines: Wheel", ["simple machines", "wheel", "science"], ["wheel simple machine", "wheels for kids", "wheels and axles"]],
  ["Simple Machines: Ramp", ["simple machines", "ramp", "science"], ["ramp simple machine", "inclined plane kids", "ramps for kids"]],
  ["Simple Machines: Pulley", ["simple machines", "pulley", "science"], ["pulley for kids", "simple machines pulley", "lift with pulley"]],
  ["Simple Machines: Screw", ["simple machines", "screw", "science"], ["screw simple machine", "simple machines screw", "tools screw"]],
  ["Simple Machines: Wedge", ["simple machines", "wedge", "science"], ["wedge simple machine", "simple machines wedge", "wedges for kids"]],
  ["Force and Motion", ["force", "motion", "science"], ["force and motion", "push pull motion", "how things move"]],
  ["Transportation", ["transportation", "wheels", "nonfiction"], ["transportation for kids", "vehicles for kids", "cars trains planes"]],
  ["Ocean Animals", ["animals", "ocean", "nonfiction"], ["ocean animals", "sea animals", "marine animals kids"]],
  ["Farm Animals", ["animals", "farm", "nonfiction"], ["farm animals", "farm books kids", "animals on a farm"]],
  ["Dinosaurs", ["animals", "dinosaurs", "nonfiction"], ["dinosaurs for kids", "dinosaur facts", "early reader dinosaurs"]],
  ["Space", ["science", "space", "nonfiction"], ["space for kids", "planets for kids", "moon stars sun"]],
  ["Feelings", ["social emotional", "feelings", "fiction"], ["feelings books kids", "emotions story", "kindness feelings"]],
  ["Healthy Bodies", ["health", "body", "nonfiction"], ["body for kids", "healthy habits", "my body book"]],
  ["Food and Nutrition", ["health", "food", "nonfiction"], ["food groups kids", "healthy food", "nutrition for kids"]],
  ["Sports and Training", ["sports", "training", "nonfiction"], ["sports training kids", "practice sports", "soccer training"]],
  ["Dance and Movement", ["dance", "movement", "fiction"], ["dance books kids", "dancing story", "movement books"]],
  ["Money and Fees", ["math vocabulary", "money", "life skills"], ["money for kids", "coins cents", "fee money"]],
  ["Real and Pretend", ["vocabulary", "fiction", "nonfiction"], ["real and pretend", "fiction nonfiction", "real or make believe"]],
  ["Float and Sink", ["science", "float", "water"], ["float and sink", "things that float", "water science kids"]],
  ["Straight and Curved Lines", ["math vocabulary", "shapes", "geometry"], ["straight curved lines", "lines and shapes", "geometry for kids"]],
  ["Whole and Parts", ["math vocabulary", "parts", "whole"], ["whole and parts", "parts of a whole", "whole half"]],
  ["Model Making", ["design", "model", "stem"], ["models for kids", "make a model", "stem model building"]],
  ["Design Thinking", ["design", "inventors", "stem"], ["design thinking kids", "design and build", "plan build test"]],
  ["Emergency Helpers", ["community", "doctor", "nonfiction"], ["emergency doctor", "ER doctor kids", "hospital helpers"]],
  ["Biography Early Readers", ["biography", "real people", "nonfiction"], ["biography for kids", "real people stories", "famous people early reader"]],
  ["Fables and Lessons", ["fiction", "fable", "lesson"], ["fables for kids", "stories with lessons", "folktales grade 1"]],
  ["Poetry and Rhymes", ["poetry", "rhyming", "fluency"], ["rhyming books", "poems for kids", "poetry grade 1"]],
  ["Dialogue and Quotation Marks", ["grammar", "dialogue", "punctuation"], ["quotation marks kids", "dialogue story", "punctuation books"]],
  ["Pronouns", ["grammar", "pronouns", "sentences"], ["pronouns for kids", "he she they", "grammar pronouns"]],
  ["Verb Tense", ["grammar", "verbs", "sentences"], ["verb tense kids", "past present verbs", "grammar verbs"]],
  ["Adjectives", ["grammar", "adjectives", "sentences"], ["adjectives for kids", "describing words", "grammar adjectives"]],
  ["Adverbs", ["grammar", "adverbs", "sentences"], ["adverbs for kids", "how words", "grammar adverbs"]]
].map(([title, tags, searchTerms]) => ({
  title,
  author: "Search mission",
  source: "Curated search entry",
  sourceUrl: "https://www.getepic.com/",
  gradeBand: "G1-G2",
  format: tags.includes("fiction") ? "Fiction" : tags.includes("phonics") ? "Decodable/Search" : "Search",
  readTime: "10-15 min",
  topicTags: tags,
  skillTags: tags,
  searchTerms
}));

const state = {
  activeLearnerId: learners[0].id,
  activeProfileId: null,
  activeRole: null,
  editingProfileId: null,
  activeSection: "today",
  practiceDomain: practiceDomains[0].id,
  activePracticeId: null,
  practiceMode: "overview",
  mediaRecorder: null,
  audioChunks: [],
  activeRecordingMimeType: "",
  readingWarmupRecordingUrl: "",
  readingWarmupRecordingActive: false,
  readingWarmupReplayActive: false,
  readingWarmupReplayAudio: null,
  readingWarmupReplayError: "",
  readingWarmupRecordingUrls: {},
  readingWarmupActiveArticleId: "",
  readingWarmupReplayArticleId: "",
  razActiveArticleId: "",
  razArticleAnalysisPending: {},
  razPdfRenderToken: 0,
  razPdfRenderError: "",
  razInventoryPopup: false,
  speakingRecordingUrls: {},
  speakingReplayErrors: {},
  retellParentChecked: {},
  retellSessionKey: "",
  heartRewardOverlay: null,
  storyRewardAudioError: "",
  activeSpeakingRecordingIndex: null,
  lastPopoverWord: "",
  activePopoverAnchor: null,
  speakingTarget: "I can read weekly words.",
  spellingIndex: 0,
  spellingFeedback: "",
  spellingFeedbackOverlay: null,
  sightFeedbackOverlay: null,
  spellingWordPulse: false,
  activePracticeStartedAt: null,
  activePracticeTimerCenterId: null,
  activeAudioKey: null,
  currentUtterance: null,
  audioSequenceTimers: [],
  parentUnlocked: false,
  childUnlocked: {},
  authProfileId: "",
  authMessage: "",
  materials: [],
  materialsLoading: false,
  materialsStatus: "",
  materialSearch: "",
  weeklySuggestions: null,
  materialBatch: null,
  weeklyPlanEditing: false,
  cloudSyncStatus: "not signed in",
  cloudSyncEmail: "",
  cloudSyncMessage: "Cloud sync starts automatically on this device.",
  cloudSyncLoading: false
};

let store = loadStore();
let cloudSaveTimer = null;
let cloudSyncHydrating = false;
let materialDbPromise = null;
let practiceTimerId = null;
let pdfLibPromise = null;
const pdfDocumentPromisesByUrl = new Map();
let xlsxLibPromise = null;
let tesseractLibPromise = null;

if (DEV_RESET_PRACTICE_ON_LOAD) {
  learners.forEach((learner) => {
    store.learners[learner.id] = clearPracticeProgress(store.learners[learner.id]);
  });
  saveStore();
}

function todayKey(date = new Date()) {
  return date.toISOString().slice(0, 10);
}

function localTodayKey(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function displayDate(date = new Date()) {
  return new Intl.DateTimeFormat("en-US", {
    weekday: "long",
    month: "long",
    day: "numeric"
  }).format(date);
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return todayKey(next);
}

function dateOffsetFrom(startDate, date) {
  const start = Date.parse(`${startDate}T00:00:00Z`);
  const target = Date.parse(`${date}T00:00:00Z`);
  return Math.floor((target - start) / 86400000);
}

function isSummerRazDate(date = todayKey()) {
  return RAZ_PUBLIC_LIBRARY_ENABLED && date >= RAZ_SUMMER_START_DATE && date <= RAZ_SUMMER_END_DATE;
}

function razLibrary() {
  return Array.isArray(window.RAZ_READING_LIBRARY) ? window.RAZ_READING_LIBRARY : [];
}

function summerRazAssignmentsForDate(date = todayKey()) {
  const publishedWeekStart = window.RAZ_READING_WEEK_START || "";
  const dayIndex = publishedWeekStart ? dateOffsetFrom(publishedWeekStart, date) : -1;
  if (dayIndex < 0 || dayIndex > 6) return [];
  return razLibrary().filter((article) => Number(article.dayIndex) === dayIndex).map((article) => ({ ...article }));
}

function defaultRazRetellProgress() {
  return {
    selectedArticleId: "",
    selectedAt: "",
    recordingDone: false,
    replayDone: false,
    parentChecked: false
  };
}

function weekKey(date = new Date()) {
  const start = new Date(date);
  const day = start.getDay() || 7;
  start.setDate(start.getDate() - day + 1);
  return todayKey(start);
}

function calendarWeekInfo(date = new Date()) {
  const target = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNumber = target.getUTCDay() || 7;
  target.setUTCDate(target.getUTCDate() + 4 - dayNumber);
  const yearStart = new Date(Date.UTC(target.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((target - yearStart) / 86400000 + 1) / 7);
  const year = target.getUTCFullYear();
  return {
    calendarYear: year,
    calendarWeek: week,
    calendarWeekLabel: `Calendar Week ${week}`,
    archiveLabel: `${year} Calendar Week ${week}`
  };
}

function planWeekInfo(plan = activePlan()) {
  if (plan.archiveLabel && plan.calendarYear && plan.calendarWeek) {
    return {
      calendarYear: plan.calendarYear,
      calendarWeek: plan.calendarWeek,
      calendarWeekLabel: plan.calendarWeekLabel || `Calendar Week ${plan.calendarWeek}`,
      archiveLabel: plan.archiveLabel
    };
  }
  return calendarWeekInfo();
}

function canonicalWeekId(year, week) {
  return `${year}-cw${String(week).padStart(2, "0")}`;
}

function learningWeekInfo(date = new Date()) {
  const nextDay = new Date(date);
  nextDay.setDate(nextDay.getDate() + 1);
  const week = calendarWeekInfo(nextDay);
  return {
    ...week,
    weekId: canonicalWeekId(week.calendarYear, week.calendarWeek)
  };
}

function activeWeekIdForDate(date = new Date()) {
  return learningWeekInfo(date).weekId;
}

function nextWeekInfoForDate(date = new Date()) {
  const next = new Date(date);
  next.setDate(next.getDate() + 7);
  return learningWeekInfo(next);
}

function planMatchesWeekId(plan, weekId) {
  return Boolean(plan?.weekId && (plan.weekId === weekId || plan.weekId.startsWith(`${weekId}-`)));
}

function isDefaultWeeklyPlanContent(plan) {
  return !plan?.sourceTitle &&
    (plan?.readingTitle === defaultWeeklyPlan.readingTitle || plan?.readingArticleTitle === defaultWeeklyPlan.readingArticleTitle) &&
    normalizeWord(plan?.masteryWords?.[0] || "") === normalizeWord(defaultWeeklyPlan.masteryWords[0]);
}

function planContentSignature(plan = {}) {
  return [
    plan.sourceTitle || "",
    plan.readingTitle || "",
    plan.readingArticleTitle || "",
    ...(plan.masteryWords ?? [])
  ].map((item) => normalizeWord(String(item))).join("|");
}

function shouldRecoverIndexedWeeklyPlan(currentPlan, indexedPlan, activeWeekId, date = new Date()) {
  if (!planMatchesWeekId(currentPlan, activeWeekId) || !planMatchesWeekId(indexedPlan, activeWeekId)) return false;
  if (indexedPlan === currentPlan) return false;
  if (planContentSignature(currentPlan) === planContentSignature(indexedPlan)) return false;
  const indexedHasUploadedContent = Boolean(indexedPlan.sourceTitle || indexedPlan.readingTitle || indexedPlan.readingArticleTitle);
  const currentLooksStale = isDefaultWeeklyPlanContent(currentPlan) || !currentPlan.sourceTitle || date.getDay() === 0;
  return indexedHasUploadedContent && currentLooksStale;
}

function builtInWeeklyPlanForWeek(weekId) {
  const seed = builtInWeeklyPlanSeeds[weekId];
  return seed ? normalizeCurrentWeeklyPlan(seed) : null;
}

function cw25StoryRewardAudioIsCurrent(listeningStories = []) {
  const stories = listeningStories.map(normalizeListeningStory).filter(Boolean);
  if (stories.length < 7) return false;
  return cw25ListeningStories().every((story) => {
    const match = stories.find((item) => Number(item.dayIndex) === Number(story.dayIndex));
    return match && String(match.audioUrl || "") === String(story.audioUrl || "") && hasPlayableAudio(match);
  });
}

function cw26StoryRewardAudioIsCurrent(listeningStories = []) {
  const stories = listeningStories.map(normalizeListeningStory).filter(Boolean);
  if (stories.length < 7) return false;
  return cw26ListeningStories().every((story) => {
    const match = stories.find((item) => Number(item.dayIndex) === Number(story.dayIndex));
    return match && String(match.audioUrl || "") === String(story.audioUrl || "") && hasPlayableAudio(match);
  });
}

function shouldUseCw26StoryRewardAudio(date = new Date()) {
  return activeWeekIdForDate(date) === "2026-cw26";
}

function ensureCw26StoryRewardAudio(date = new Date()) {
  if (!shouldUseCw26StoryRewardAudio(date)) return false;
  const stories = cw26ListeningStories();
  if (stories.length < 7) return false;
  let changed = false;
  const patchPlan = (plan) => {
    if (!plan || !planMatchesWeekId(plan, "2026-cw26")) return plan;
    if (cw26StoryRewardAudioIsCurrent(plan.listeningStories ?? [])) return plan;
    changed = true;
    return {
      ...plan,
      listeningStories: stories,
      updatedAt: new Date().toISOString()
    };
  };
  store.weeklyPlan = patchPlan(store.weeklyPlan);
  store.pendingWeeklyPlan = patchPlan(store.pendingWeeklyPlan);
  if (store.weeklyPlansByWeekId?.["2026-cw26"]) {
    store.weeklyPlansByWeekId["2026-cw26"] = patchPlan(store.weeklyPlansByWeekId["2026-cw26"]);
  }
  if (store.cw26StoryRewardAudioFixVersion !== CW26_STORY_REWARD_AUDIO_FIX_VERSION) {
    store.cw26StoryRewardAudioFixVersion = CW26_STORY_REWARD_AUDIO_FIX_VERSION;
    changed = true;
  }
  if (changed) saveStore();
  return changed;
}

function shouldUseCw25StoryRewardAudio(plan = activePlan(), date = new Date()) {
  return planMatchesWeekId(plan, "2026-cw25") && localTodayKey(date) >= CW25_STORY_REWARD_AUDIO_FIX_START;
}

function ensureCw25StoryRewardAudio(date = new Date()) {
  if (localTodayKey(date) < CW25_STORY_REWARD_AUDIO_FIX_START) return false;
  const stories = cw25ListeningStories();
  if (stories.length < 7) return false;
  let changed = false;
  const patchPlan = (plan) => {
    if (!plan || !planMatchesWeekId(plan, "2026-cw25")) return plan;
    if (cw25StoryRewardAudioIsCurrent(plan.listeningStories ?? [])) return plan;
    changed = true;
    return {
      ...plan,
      listeningStories: stories,
      updatedAt: new Date().toISOString()
    };
  };
  store.weeklyPlan = patchPlan(store.weeklyPlan);
  store.pendingWeeklyPlan = patchPlan(store.pendingWeeklyPlan);
  if (store.weeklyPlansByWeekId?.["2026-cw25"]) {
    store.weeklyPlansByWeekId["2026-cw25"] = patchPlan(store.weeklyPlansByWeekId["2026-cw25"]);
  }
  if (changed) saveStore();
  return changed;
}

function patchCw25WeeklyPlan(plan) {
  if (!plan || !planMatchesWeekId(plan, "2026-cw25")) return plan;
  const fixed = normalizeCurrentWeeklyPlan(builtInWeeklyPlanSeeds["2026-cw25"]);
  return {
    ...plan,
    readingTitle: fixed.readingTitle,
    readingArticleTitle: fixed.readingArticleTitle,
    readingArticleGenre: fixed.readingArticleGenre,
    readingArticleSummary: fixed.readingArticleSummary,
    readingArticleContentStatus: fixed.readingArticleContentStatus,
    readingArticleEvidence: fixed.readingArticleEvidence,
    readingArticleCharacters: [],
    readingArticleSetting: "",
    readingArticleProblem: "",
    readingArticleKeyEvents: [],
    readingArticleEnding: "",
    readingArticleBigIdea: "Poems can use repeated words, rhythm, rhyme, and images to show feelings and ideas.",
    phonicsFocus: fixed.phonicsFocus,
    grammarFocus: fixed.grammarFocus,
    uoiTopic: fixed.uoiTopic,
    masteryWords: [...fixed.masteryWords],
    sightWords: [...fixed.sightWords],
    highFrequencyWords: [...fixed.highFrequencyWords],
    grammarPracticeWords: [...fixed.grammarPracticeWords],
    readingSkills: [...fixed.readingSkills],
    uiVocabulary: fixed.uiVocabulary.map((item) => ({ ...item })),
    englishVocabulary: fixed.englishVocabulary.map((item) => ({ ...item })),
    spelling: {
      skill: fixed.spelling.skill,
      basic: [...fixed.spelling.basic],
      review: [...fixed.spelling.review],
      challenge: [...fixed.spelling.challenge]
    },
    weeklyVocabulary: fixed.weeklyVocabulary.map((item) => ({ ...item })),
    readingSelections: [...fixed.readingSelections],
    poemSelections: fixed.poemSelections.map((item) => ({ ...item })),
    jumpRopeReader: fixed.jumpRopeReader,
    weeklyLearningOutcomes: [...fixed.weeklyLearningOutcomes],
    homeLearningSuggestions: [...fixed.homeLearningSuggestions],
    epicIdeas: [...fixed.epicIdeas],
    sourceTitle: fixed.sourceTitle,
    listeningStories: cw25ListeningStories(),
    updatedAt: new Date().toISOString()
  };
}

function ensureCw25ManualWeeklyPlanFix() {
  if (store.cw25ManualPlanFixVersion === CW25_MANUAL_PLAN_FIX_VERSION) return false;
  let changed = false;
  const patch = (plan) => {
    const patched = patchCw25WeeklyPlan(plan);
    if (patched !== plan) changed = true;
    return patched;
  };
  store.weeklyPlan = patch(store.weeklyPlan);
  store.pendingWeeklyPlan = patch(store.pendingWeeklyPlan);
  if (store.weeklyPlansByWeekId?.["2026-cw25"]) {
    store.weeklyPlansByWeekId["2026-cw25"] = patch(store.weeklyPlansByWeekId["2026-cw25"]);
  }
  store.cw25ManualPlanFixVersion = CW25_MANUAL_PLAN_FIX_VERSION;
  saveStore();
  return changed;
}

function applyWeekMetadata(plan, weekInfo, status = plan.status || "active") {
  return {
    ...normalizeCurrentWeeklyPlan(plan),
    weekId: weekInfo.weekId,
    calendarYear: weekInfo.calendarYear,
    calendarWeek: weekInfo.calendarWeek,
    calendarWeekLabel: weekInfo.calendarWeekLabel,
    archiveLabel: weekInfo.archiveLabel,
    activationDate: plan.activationDate || weekInfo.activationDate || "",
    status
  };
}

function isFridayOrSaturday(date = new Date()) {
  return [5, 6].includes(date.getDay());
}

function resetWeekScopedProgress(progress = getDefaultProgress()) {
  return {
    ...progress,
    completedDays: [],
    centerCompletions: {},
    adventureProgress: {},
    practiceStats: {},
    spellingSessions: {},
    sightSessions: {},
    readingQuizSessions: {},
    grammarSessions: {},
    grammarSkillProfile: {},
    weeklyExposure: {},
    storyWordExposure: {},
    readingQuizAttempts: [],
    grammarPracticeAttempts: [],
    speakingAttempts: []
  };
}

function clearWeekScopedPracticeForAllLearners() {
  learners.forEach((learner) => {
    store.learners[learner.id] = resetWeekScopedProgress(store.learners[learner.id]);
  });
}

function clearTrainingProgressForAllLearners() {
  learners.forEach((learner) => {
    store.learners[learner.id] = {
      ...resetWeekScopedProgress(store.learners[learner.id]),
      reviewQueue: {}
    };
  });
}

function resetLocalTrainingPlans() {
  const weekInfo = learningWeekInfo();
  const resetPlan = applyWeekMetadata({ ...defaultWeeklyPlan }, weekInfo, "active");
  store.weeklyPlan = resetPlan;
  store.currentWeekId = resetPlan.weekId;
  store.lastActivatedWeekId = resetPlan.weekId;
  store.pendingWeeklyPlan = null;
  store.weeklyPlansByWeekId = { [weekInfo.weekId]: resetPlan };
  clearTrainingProgressForAllLearners();
}

function ensureWeeklyPlanIndex() {
  if (!store.weeklyPlansByWeekId || typeof store.weeklyPlansByWeekId !== "object") store.weeklyPlansByWeekId = {};
  const plan = normalizeCurrentWeeklyPlan(store.weeklyPlan);
  const currentInfo = plan.calendarYear && plan.calendarWeek
    ? { ...planWeekInfo(plan), weekId: canonicalWeekId(plan.calendarYear, plan.calendarWeek) }
    : learningWeekInfo();
  const weekId = plan.weekId?.match(/^\d{4}-cw\d{2}/)?.[0] || currentInfo.weekId;
  store.weeklyPlan = { ...plan, weekId: plan.weekId || weekId };
  store.currentWeekId = store.weeklyPlan.weekId;
  store.weeklyPlansByWeekId[weekId] = store.weeklyPlansByWeekId[weekId] || store.weeklyPlan;
}

function activateWeeklyPlanForToday(date = new Date(), { render = false } = {}) {
  ensureWeeklyPlanIndex();
  const activeWeekId = activeWeekIdForDate(date);
  const pending = store.pendingWeeklyPlan;
  const indexed = store.weeklyPlansByWeekId?.[activeWeekId];
  const builtInPlan = builtInWeeklyPlanForWeek(activeWeekId);
  const shouldPromotePending = Boolean(pending) && (planMatchesWeekId(pending, activeWeekId) || pending.status === "pending" || date.getDay() === 0);
  const shouldUseIndexed = Boolean(indexed) && shouldRecoverIndexedWeeklyPlan(store.weeklyPlan, indexed, activeWeekId, date);
  const shouldUseBuiltIn = Boolean(builtInPlan) &&
    !shouldPromotePending &&
    !shouldUseIndexed &&
    planContentSignature(store.weeklyPlan) !== planContentSignature(builtInPlan) &&
    (isDefaultWeeklyPlanContent(store.weeklyPlan) || !planMatchesWeekId(store.weeklyPlan, activeWeekId) || date.getDay() === 0);
  if (planMatchesWeekId(store.weeklyPlan, activeWeekId) && !shouldPromotePending && !shouldUseIndexed && !shouldUseBuiltIn) return false;
  const nextPlan = shouldPromotePending ? pending : shouldUseIndexed ? indexed : shouldUseBuiltIn ? builtInPlan : indexed;
  if (!nextPlan) return false;

  archiveCurrentWeek("Auto-archived before weekly rollover.", { render: false });
  const activePlan = applyWeekMetadata(nextPlan, learningWeekInfo(date), "active");
  store.weeklyPlan = activePlan;
  store.currentWeekId = activePlan.weekId;
  store.lastActivatedWeekId = activePlan.weekId;
  store.weeklyPlansByWeekId[activeWeekId] = activePlan;
  if (shouldPromotePending || planMatchesWeekId(store.pendingWeeklyPlan, activeWeekId)) store.pendingWeeklyPlan = null;
  clearWeekScopedPracticeForAllLearners();
  seedReviewQueueForPlan(activePlan, activePlan.sourceTitle || activePlan.readingArticleTitle || "weekly words");
  saveStore();
  if (render) renderAll();
  return true;
}

function normalizeWord(word) {
  return word.toLowerCase().replace(/^[^a-z']+|[^a-z']+$/g, "");
}

function displayWord(word) {
  const normalized = normalizeWord(word);
  const weeklyWord = activePlan().masteryWords?.find((item) => normalizeWord(item) === normalized);
  const sightWord = activePlan().sightWords?.find((item) => normalizeWord(item) === normalized);
  return weeklyWord || sightWord || word;
}

function uniqueWords(words) {
  return [...new Set(words.map((word) => word.trim()).filter(Boolean))];
}

function splitEditableList(value) {
  return uniqueWords(
    String(value || "")
      .split(/,|\n|;/)
      .map((item) => item.trim().replace(/\s+/g, " "))
      .filter(Boolean)
  );
}

function splitEditableLines(value) {
  return uniqueWords(
    String(value || "")
      .split(/\n|\|/)
      .map((item) => item.trim().replace(/\s+/g, " "))
      .filter(Boolean)
  );
}

function splitRetellEvidence(value) {
  return uniqueWords(
    String(value || "")
      .split(/\n|\||(?<=[.!?])\s+/)
      .map((item) => item.trim().replace(/\s+/g, " "))
      .filter((item) => item.length >= 6)
  );
}

function normalizePoemSelection(item = {}) {
  const raw = typeof item === "string" ? { title: item } : item;
  if (!raw || typeof raw !== "object") return null;
  const title = String(raw.title || "").trim().replace(/\s+/g, " ");
  if (!title) return null;
  return {
    title,
    author: String(raw.author || "").trim().replace(/\s+/g, " "),
    status: String(raw.status || "").trim().replace(/\s+/g, " ") || "parent-provided",
    evidence: splitRetellEvidence(raw.evidence || raw.summary || "").join(" "),
    focus: String(raw.focus || "").trim().replace(/\s+/g, " "),
    sourceUrl: String(raw.sourceUrl || "").trim()
  };
}

function normalizePoemSelections(items = []) {
  const byTitle = new Map();
  (Array.isArray(items) ? items : []).forEach((item) => {
    const poem = normalizePoemSelection(item);
    if (!poem) return;
    const key = normalizeWord(poem.title);
    if (!byTitle.has(key)) byTitle.set(key, poem);
  });
  return [...byTitle.values()];
}

function formatPoemSelectionsForReview(items = []) {
  return normalizePoemSelections(items)
    .map((poem) => [poem.title, poem.evidence].filter(Boolean).join(": "))
    .join("\n");
}

function parsePoemSelectionsFromReview(value) {
  return uniqueWords(
    String(value || "")
      .split(/\n/)
      .map((item) => item.trim())
      .filter(Boolean)
  ).map((line) => {
    const separatorIndex = line.indexOf(":");
    if (separatorIndex < 0) return { title: line, evidence: "" };
    return {
      title: line.slice(0, separatorIndex).trim(),
      evidence: line.slice(separatorIndex + 1).trim(),
      status: "parent-edited"
    };
  });
}

function normalizeVocabularyEntry(item, category = "weekly") {
  const raw = typeof item === "string" ? { word: item } : item;
  if (!raw || typeof raw !== "object") return null;
  const word = String(raw.word || "").trim();
  if (!normalizeWord(word)) return null;
  const key = normalizeWord(word);
  return {
    word,
    translation: raw.translation || "",
    definition: raw.definition || raw.translation || wordHelp[key]?.[0] || "a word from this week's learning",
    example: raw.example || wordHelp[key]?.[1] || `Use ${word} in a complete sentence.`,
    category: raw.category || category,
    source: raw.source || "weekly plan"
  };
}

function normalizeVocabularyEntries(items = [], category = "weekly") {
  const byWord = new Map();
  (Array.isArray(items) ? items : []).forEach((item) => {
    const entry = normalizeVocabularyEntry(item, category);
    if (!entry) return;
    const key = normalizeWord(entry.word);
    if (!byWord.has(key)) byWord.set(key, entry);
  });
  return [...byWord.values()];
}

function wordsToVocabularyEntries(words = [], category = "weekly") {
  return normalizeVocabularyEntries((Array.isArray(words) ? words : []).map((word) => ({ word })), category);
}

function cw26VocabularyEntries() {
  const cw25 = builtInWeeklyPlanSeeds["2026-cw25"] ?? {};
  return {
    uiVocabulary: normalizeVocabularyEntries([
      ...(cw25.uiVocabulary ?? []),
      { word: "share", translation: "分享", definition: "to give, show, or use something with other people", example: "Scientists share their ideas with the class." },
      { word: "proud", translation: "自豪的", definition: "feeling pleased about something you did or learned", example: "I feel proud when I finish a hard project." },
      { word: "scientist", translation: "科学家", definition: "a person who studies the world by asking questions and testing ideas", example: "A scientist observes carefully and records what happens." }
    ], "uiVocabulary"),
    englishVocabulary: normalizeVocabularyEntries([
      ...(cw25.englishVocabulary ?? []),
      { word: "learn", translation: "学习", definition: "to gain knowledge or a new skill", example: "We learn by asking questions and trying new ideas." },
      { word: "grow", translation: "成长；生长", definition: "to become bigger, older, or more skilled", example: "Plants grow when they have light and water." },
      { word: "help", translation: "帮助", definition: "to make something easier or better for someone", example: "Friends help each other solve problems." },
      { word: "try", translation: "尝试", definition: "to make an effort to do something", example: "I will try a new way if the first plan does not work." },
      { word: "fun", translation: "有趣；乐趣", definition: "something enjoyable that makes you happy", example: "The science activity is fun because we can test our ideas." },
      { word: "hard", translation: "困难的；努力地", definition: "difficult to do or needing a lot of effort", example: "The puzzle is hard, but I keep trying." },
      { word: "easy", translation: "容易的", definition: "not difficult to do", example: "The first step is easy when we work together." },
      { word: "remember", translation: "记得", definition: "to keep or bring something back into your mind", example: "Remember to share what you learned." },
      { word: "year", translation: "年", definition: "a period of twelve months", example: "This year, we learned how scientists solve problems." },
      { word: "next", translation: "下一个；接下来", definition: "coming immediately after something", example: "Next, we test the design and record what happens." }
    ], "englishVocabulary")
  };
}

function shouldUseCw26Vocabulary(date = new Date()) {
  return activeWeekIdForDate(date) === "2026-cw26";
}

function shouldApplyCw26VocabularyToPlan(plan, date = new Date()) {
  return planMatchesWeekId(plan, "2026-cw26") || (shouldUseCw26Vocabulary(date) && plan === activePlan());
}

function vocabularyEntriesForPlan(plan = activePlan(), date = new Date()) {
  if (shouldApplyCw26VocabularyToPlan(plan, date)) return cw26VocabularyEntries();
  return {
    uiVocabulary: normalizeVocabularyEntries(plan.uiVocabulary ?? [], "uiVocabulary"),
    englishVocabulary: normalizeVocabularyEntries(plan.englishVocabulary ?? [], "englishVocabulary")
  };
}

function cw26VocabularyIsCurrent(plan) {
  if (!plan || !planMatchesWeekId(plan, "2026-cw26")) return false;
  const expected = cw26VocabularyEntries();
  const signature = (items) => normalizeVocabularyEntries(items).map((item) => [
    normalizeWord(item.word),
    item.translation,
    item.definition,
    item.example
  ]);
  return JSON.stringify(signature(plan.uiVocabulary ?? [])) === JSON.stringify(signature(expected.uiVocabulary)) &&
    JSON.stringify(signature(plan.englishVocabulary ?? [])) === JSON.stringify(signature(expected.englishVocabulary));
}

function ensureCw26Vocabulary(date = new Date()) {
  if (!shouldUseCw26Vocabulary(date)) return false;
  const vocabulary = cw26VocabularyEntries();
  let changed = false;
  const patchPlan = (plan) => {
    if (!plan || !planMatchesWeekId(plan, "2026-cw26") || cw26VocabularyIsCurrent(plan)) return plan;
    changed = true;
    return {
      ...plan,
      uiVocabulary: vocabulary.uiVocabulary.map((item) => ({ ...item })),
      englishVocabulary: vocabulary.englishVocabulary.map((item) => ({ ...item })),
      updatedAt: new Date().toISOString()
    };
  };
  store.weeklyPlan = patchPlan(store.weeklyPlan);
  store.pendingWeeklyPlan = patchPlan(store.pendingWeeklyPlan);
  if (store.weeklyPlansByWeekId?.["2026-cw26"]) {
    store.weeklyPlansByWeekId["2026-cw26"] = patchPlan(store.weeklyPlansByWeekId["2026-cw26"]);
  }
  if (store.cw26VocabularyFixVersion !== CW26_VOCABULARY_FIX_VERSION) {
    store.cw26VocabularyFixVersion = CW26_VOCABULARY_FIX_VERSION;
    changed = true;
  }
  if (changed) saveStore();
  return changed;
}

function cw24UiVocabularyEntries() {
  return normalizeVocabularyEntries(builtInWeeklyPlanSeeds["2026-cw24"]?.uiVocabulary ?? [], "uiVocabulary");
}

function cw24EnglishVocabularyEntries() {
  return normalizeVocabularyEntries(builtInWeeklyPlanSeeds["2026-cw24"]?.englishVocabulary ?? [], "englishVocabulary");
}

function emptySpellingGroups() {
  return { skill: "", basic: [], review: [], challenge: [] };
}

function normalizeSpellingGroups(spelling = {}) {
  const source = spelling && typeof spelling === "object" ? spelling : {};
  return {
    skill: source.skill || "",
    basic: uniqueWords(source.basic ?? []),
    review: uniqueWords(source.review ?? []),
    challenge: uniqueWords(source.challenge ?? [])
  };
}

function spellingWordsFromGroups(spelling = emptySpellingGroups()) {
  return uniqueWords([...(spelling.basic ?? []), ...(spelling.review ?? []), ...(spelling.challenge ?? [])]);
}

function spellingPriorityWords(plan = activePlan()) {
  const vocabulary = vocabularyEntriesForPlan(plan);
  return uniqueWords([
    ...vocabulary.englishVocabulary.map((item) => item.word),
    ...spellingWordsFromGroups(plan.spelling),
    ...vocabulary.uiVocabulary.map((item) => item.word)
  ].map(displayWord)).filter(Boolean);
}

function planWeeklyNewWords(plan = activePlan()) {
  return spellingPriorityWords(plan);
}

function escapeHTML(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function getDefaultProgress() {
  return {
    completedDays: [],
    centerCompletions: {},
    adventureProgress: {},
    wordStats: {},
    reviewQueue: {},
    practiceStats: {},
    spellingSessions: {},
    sightSessions: {},
    readingQuizSessions: {},
    grammarSessions: {},
    heartRewardDays: [],
    grammarSkillProfile: {},
    sightWordProfile: {},
    weeklyExposure: {},
    storyWordExposure: {},
    wrongWordProfile: {},
    epicLog: [],
    examScores: [],
    speakingAttempts: [],
    readingQuizAttempts: [],
    grammarPracticeAttempts: [],
    weakSkills: ["long vowels", "syllable division", "complete sentence answers"]
  };
}

function clearPracticeProgress(progress = getDefaultProgress()) {
  return {
    ...progress,
    completedDays: [],
    centerCompletions: {},
    adventureProgress: {},
    wordStats: {},
    reviewQueue: {},
    practiceStats: {},
    spellingSessions: {},
    sightSessions: {},
    readingQuizSessions: {},
    grammarSessions: {},
    heartRewardDays: [],
    grammarSkillProfile: {},
    sightWordProfile: {},
    weeklyExposure: {},
    storyWordExposure: {},
    wrongWordProfile: {},
    epicLog: [],
    examScores: [],
    speakingAttempts: [],
    readingQuizAttempts: [],
    grammarPracticeAttempts: [],
    weakSkills: ["long vowels", "syllable division", "complete sentence answers"]
  };
}

function normalizeCurrentWeeklyPlan(rawPlan = {}) {
  const plan = rawPlan && typeof rawPlan === "object" ? rawPlan : {};
  const weeklySequence = plan.weeklySequence && typeof plan.weeklySequence === "object" ? plan.weeklySequence : null;
  const sequenceSpelling = normalizeSpellingGroups(weeklySequence?.spelling ?? {});
  const planSpelling = normalizeSpellingGroups(plan.spelling ?? {});
  const spelling = normalizeSpellingGroups({
    skill: planSpelling.skill || sequenceSpelling.skill,
    basic: uniqueWords([...(planSpelling.basic ?? []), ...(sequenceSpelling.basic ?? [])]),
    review: uniqueWords([...(planSpelling.review ?? []), ...(sequenceSpelling.review ?? [])]),
    challenge: uniqueWords([...(planSpelling.challenge ?? []), ...(sequenceSpelling.challenge ?? [])])
  });
  const isCw24Plan = plan.weekId === "2026-cw24" || plan.calendarWeek === 24 || /joaquin'?s zoo/i.test(`${plan.readingTitle || ""} ${plan.readingArticleTitle || ""}`);
  const uiVocabulary = normalizeVocabularyEntries(
    plan.uiVocabulary?.length ? plan.uiVocabulary : isCw24Plan ? cw24UiVocabularyEntries() : [],
    "uiVocabulary"
  );
  const englishFromSequence = uniqueWords([
    ...(weeklySequence?.vocabulary?.oralPower ?? []),
    ...(weeklySequence?.vocabulary?.power ?? [])
  ]);
  const englishVocabulary = normalizeVocabularyEntries(
    plan.englishVocabulary?.length
      ? plan.englishVocabulary
      : isCw24Plan
        ? cw24EnglishVocabularyEntries()
        : wordsToVocabularyEntries(englishFromSequence, "englishVocabulary"),
    "englishVocabulary"
  );
  const highFrequencyWords = uniqueWords(
    plan.highFrequencyWords?.length
      ? plan.highFrequencyWords
      : weeklySequence?.highFrequencyWords?.length
        ? weeklySequence.highFrequencyWords
        : isCw24Plan
          ? builtInWeeklyPlanSeeds["2026-cw24"].highFrequencyWords
          : []
  );

  return {
    ...defaultWeeklyPlan,
    ...plan,
    masteryWords: Array.isArray(plan.masteryWords) ? [...plan.masteryWords] : [...defaultWeeklyPlan.masteryWords],
    sightWords: Array.isArray(plan.sightWords) ? [...plan.sightWords] : [...defaultWeeklyPlan.sightWords],
    grammarPracticeWords: Array.isArray(plan.grammarPracticeWords) ? [...plan.grammarPracticeWords] : [...defaultWeeklyPlan.grammarPracticeWords],
    readingSkills: Array.isArray(plan.readingSkills) ? [...plan.readingSkills] : [...defaultWeeklyPlan.readingSkills],
    uiVocabulary,
    englishVocabulary,
    spelling,
    highFrequencyWords,
    weeklySequence: weeklySequence ? { ...weeklySequence } : plan.weeklySequence,
    weeklyVocabulary: Array.isArray(plan.weeklyVocabulary) ? plan.weeklyVocabulary.map((item) => ({ ...item })) : defaultWeeklyPlan.weeklyVocabulary.map((item) => ({ ...item })),
    readingSelections: Array.isArray(plan.readingSelections) ? [...plan.readingSelections] : Array.isArray(defaultWeeklyPlan.readingSelections) ? [...defaultWeeklyPlan.readingSelections] : [],
    poemSelections: normalizePoemSelections(plan.poemSelections ?? []),
    listeningStories: Array.isArray(plan.listeningStories) ? plan.listeningStories.map(normalizeListeningStory).filter(Boolean) : [],
    epicIdeas: Array.isArray(plan.epicIdeas) ? [...plan.epicIdeas] : [...defaultWeeklyPlan.epicIdeas]
  };
}

function loadStore() {
  const fallback = {
    currentWeekId: CURRENT_WEEK_ID,
    weeklyPlan: { ...defaultWeeklyPlan },
    weeklyPlansByWeekId: { [CURRENT_WEEK_ID]: { ...defaultWeeklyPlan } },
    pendingWeeklyPlan: null,
    lastActivatedWeekId: CURRENT_WEEK_ID,
    materialTagMigrationVersion: "",
    testerDailyResetKey: "",
    profileNames: Object.fromEntries(profiles.map((profile) => [profile.id, profile.defaultName])),
    parentBooks: [],
    weeklyArchives: [],
    wordTestBank: {},
    razArticleDetails: {},
    parentPasswordHash: DEFAULT_PARENT_PASSWORD_HASH,
    learners: Object.fromEntries(learners.map((learner) => [learner.id, getDefaultProgress()]))
  };

  try {
    const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (!parsed) return fallback;
    const weeklyPlan = normalizeCurrentWeeklyPlan(parsed.weeklyPlan);
    return {
      currentWeekId: parsed.currentWeekId || weeklyPlan.weekId || CURRENT_WEEK_ID,
      weeklyPlan,
      weeklyPlansByWeekId: parsed.weeklyPlansByWeekId && typeof parsed.weeklyPlansByWeekId === "object" ? parsed.weeklyPlansByWeekId : { [weeklyPlan.weekId || CURRENT_WEEK_ID]: weeklyPlan },
      pendingWeeklyPlan: parsed.pendingWeeklyPlan ? normalizeCurrentWeeklyPlan(parsed.pendingWeeklyPlan) : null,
      lastActivatedWeekId: parsed.lastActivatedWeekId || parsed.currentWeekId || weeklyPlan.weekId || CURRENT_WEEK_ID,
      materialTagMigrationVersion: parsed.materialTagMigrationVersion || "",
      testerDailyResetKey: parsed.testerDailyResetKey || "",
      profileNames: { ...fallback.profileNames, ...(parsed.profileNames ?? {}) },
      parentBooks: Array.isArray(parsed.parentBooks) ? parsed.parentBooks : [],
      weeklyArchives: Array.isArray(parsed.weeklyArchives) ? parsed.weeklyArchives : [],
      wordTestBank: parsed.wordTestBank && typeof parsed.wordTestBank === "object" ? parsed.wordTestBank : {},
      razArticleDetails: parsed.razArticleDetails && typeof parsed.razArticleDetails === "object" ? parsed.razArticleDetails : {},
      parentPasswordHash: parsed.parentPasswordHash || DEFAULT_PARENT_PASSWORD_HASH,
      learners: Object.fromEntries(
        learners.map((learner) => {
          const progress = { ...getDefaultProgress(), ...(parsed.learners?.[learner.id] ?? {}) };
          return [learner.id, progress];
        })
      )
    };
  } catch {
    return fallback;
  }
}

function saveStore() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
  scheduleCloudSave();
}

function cloudSyncAvailable() {
  return Boolean(window.ReadingBridgeCloud);
}

function updateCloudSyncStatus(status, message = "", email = state.cloudSyncEmail) {
  state.cloudSyncStatus = status;
  state.cloudSyncMessage = message || status;
  state.cloudSyncEmail = email || "";
  const signedIn = cloudSyncAvailable() && window.ReadingBridgeCloud.isSignedIn();
  const statusTarget = document.querySelector("#cloudSyncStatus");
  const messageTarget = document.querySelector("#cloudSyncMessage");
  const syncButton = document.querySelector("[data-cloud-sync-now]");
  if (statusTarget) statusTarget.textContent = state.cloudSyncStatus;
  if (messageTarget) messageTarget.textContent = state.cloudSyncMessage;
  if (syncButton) syncButton.disabled = !signedIn;
}

function scheduleCloudSave() {
  if (cloudSyncHydrating || !cloudSyncAvailable() || !window.ReadingBridgeCloud.isSignedIn()) return;
  window.clearTimeout(cloudSaveTimer);
  cloudSaveTimer = window.setTimeout(syncStoreToCloud, 1200);
}

async function syncStoreToCloud() {
  if (!cloudSyncAvailable() || !window.ReadingBridgeCloud.isSignedIn()) return false;
  try {
    updateCloudSyncStatus("syncing", "Saving progress to cloud...");
    await window.ReadingBridgeCloud.saveCloudStore(cloudStorePayload());
    updateCloudSyncStatus("synced", "Cloud sync on. Progress is up to date.");
    return true;
  } catch (error) {
    updateCloudSyncStatus("sync error", error?.message || "Cloud save failed. Local progress is still saved.");
    return false;
  }
}

function cloudStorePayload() {
  const payload = JSON.parse(JSON.stringify(store));
  if (payload.learners) delete payload.learners.tester;
  return payload;
}

function planLearningWordCount(plan = {}) {
  return [
    ...(plan.masteryWords ?? []),
    ...(plan.sightWords ?? []),
    ...(plan.highFrequencyWords ?? []),
    ...(plan.grammarPracticeWords ?? []),
    ...(plan.spelling?.basic ?? []),
    ...(plan.spelling?.review ?? []),
    ...(plan.spelling?.challenge ?? []),
    ...(plan.uiVocabulary ?? []).map((item) => typeof item === "string" ? item : item?.word),
    ...(plan.englishVocabulary ?? []).map((item) => typeof item === "string" ? item : item?.word),
    ...(plan.weeklyVocabulary ?? []).map((item) => typeof item === "string" ? item : item?.word)
  ].filter(Boolean).length;
}

function planUpdatedTime(plan = {}) {
  const value = Date.parse(plan.updatedAt || plan.createdAt || plan.activationDate || "");
  return Number.isFinite(value) ? value : 0;
}

function weeklyPlanMergeScore(plan = {}, targetWeekId = activeWeekIdForDate()) {
  if (!plan || typeof plan !== "object") return -1;
  const normalized = normalizeCurrentWeeklyPlan(plan);
  let score = 0;
  if (planMatchesWeekId(normalized, targetWeekId)) score += 10000;
  if (normalized.sourceTitle) score += 1200;
  if (!isDefaultWeeklyPlanContent(normalized)) score += 900;
  score += Math.min(planLearningWordCount(normalized), 40) * 10;
  score += Math.min(planUpdatedTime(normalized) / 1000000000000, 9);
  return score;
}

function chooseWeeklyPlanForCloudMerge(localPlan, cloudPlan, targetWeekId = activeWeekIdForDate()) {
  if (!localPlan) return cloudPlan ? normalizeCurrentWeeklyPlan(cloudPlan) : localPlan;
  if (!cloudPlan) return normalizeCurrentWeeklyPlan(localPlan);
  const local = normalizeCurrentWeeklyPlan(localPlan);
  const cloud = normalizeCurrentWeeklyPlan(cloudPlan);
  const localScore = weeklyPlanMergeScore(local, targetWeekId);
  const cloudScore = weeklyPlanMergeScore(cloud, targetWeekId);
  if (cloudScore > localScore) return cloud;
  if (localScore > cloudScore) return local;
  return planUpdatedTime(cloud) > planUpdatedTime(local) ? cloud : local;
}

function mergeWeeklyPlansByWeekId(localPlans = {}, cloudPlans = {}) {
  const merged = { ...(cloudPlans ?? {}) };
  const ids = uniqueWords([...Object.keys(cloudPlans ?? {}), ...Object.keys(localPlans ?? {})]);
  ids.forEach((weekId) => {
    merged[weekId] = chooseWeeklyPlanForCloudMerge(localPlans?.[weekId], cloudPlans?.[weekId], weekId);
  });
  return merged;
}

function mergeUniqueList(local = [], cloud = []) {
  return uniqueWords([...(Array.isArray(cloud) ? cloud : []), ...(Array.isArray(local) ? local : [])]);
}

function mergeRazReadingArticles(local = {}, cloud = {}) {
  const merged = { ...cloud, ...local };
  Object.keys({ ...cloud, ...local }).forEach((articleId) => {
    const localArticle = local?.[articleId] ?? {};
    const cloudArticle = cloud?.[articleId] ?? {};
    merged[articleId] = {
      ...cloudArticle,
      ...localArticle,
      recordingDone: Boolean(localArticle.recordingDone || cloudArticle.recordingDone),
      replayDone: Boolean(localArticle.replayDone || cloudArticle.replayDone)
    };
  });
  return merged;
}

function mergeRazRetell(local = {}, cloud = {}) {
  const localSelectedAt = Date.parse(local.selectedAt || "") || 0;
  const cloudSelectedAt = Date.parse(cloud.selectedAt || "") || 0;
  const selected = localSelectedAt >= cloudSelectedAt ? local : cloud;
  const sameArticle = local.selectedArticleId && local.selectedArticleId === cloud.selectedArticleId;
  return {
    ...defaultRazRetellProgress(),
    ...cloud,
    ...local,
    selectedArticleId: selected.selectedArticleId || local.selectedArticleId || cloud.selectedArticleId || "",
    selectedAt: selected.selectedAt || local.selectedAt || cloud.selectedAt || "",
    recordingDone: sameArticle ? Boolean(local.recordingDone || cloud.recordingDone) : Boolean(selected.recordingDone),
    replayDone: sameArticle ? Boolean(local.replayDone || cloud.replayDone) : Boolean(selected.replayDone),
    parentChecked: sameArticle ? Boolean(local.parentChecked || cloud.parentChecked) : Boolean(selected.parentChecked)
  };
}

function mergeAdventureDayProgress(localDay = {}, cloudDay = {}) {
  return {
    ...cloudDay,
    ...localDay,
    completedLevels: mergeUniqueList(localDay.completedLevels, cloudDay.completedLevels),
    wordPowerCompletedWords: mergeUniqueList(localDay.wordPowerCompletedWords, cloudDay.wordPowerCompletedWords),
    readingWarmupRecordingDone: Boolean(localDay.readingWarmupRecordingDone || cloudDay.readingWarmupRecordingDone),
    readingWarmupReplayDone: Boolean(localDay.readingWarmupReplayDone || cloudDay.readingWarmupReplayDone),
    readingWarmupSkipped: Boolean(localDay.readingWarmupSkipped || cloudDay.readingWarmupSkipped),
    razReadingAssignments: Array.isArray(localDay.razReadingAssignments) && localDay.razReadingAssignments.length
      ? [...localDay.razReadingAssignments]
      : [...(cloudDay.razReadingAssignments ?? [])],
    razReadingArticles: mergeRazReadingArticles(localDay.razReadingArticles, cloudDay.razReadingArticles),
    razRetell: mergeRazRetell(localDay.razRetell, cloudDay.razRetell),
    storyRewardComplete: Boolean(localDay.storyRewardComplete || cloudDay.storyRewardComplete)
  };
}

function mergeAdventureProgress(local = {}, cloud = {}) {
  const merged = { ...cloud, ...local };
  Object.keys({ ...cloud, ...local }).forEach((date) => {
    merged[date] = mergeAdventureDayProgress(local[date], cloud[date]);
  });
  return merged;
}

function mergeLearnerProgress(localProgress = getDefaultProgress(), cloudProgress = getDefaultProgress()) {
  return {
    ...cloudProgress,
    ...localProgress,
    heartRewardDays: mergeUniqueList(localProgress.heartRewardDays, cloudProgress.heartRewardDays),
    completedDays: mergeUniqueList(localProgress.completedDays, cloudProgress.completedDays),
    weakSkills: mergeUniqueList(localProgress.weakSkills, cloudProgress.weakSkills),
    adventureProgress: mergeAdventureProgress(localProgress.adventureProgress, cloudProgress.adventureProgress),
    centerCompletions: {
      ...(cloudProgress.centerCompletions ?? {}),
      ...(localProgress.centerCompletions ?? {})
    },
    practiceStats: {
      ...(cloudProgress.practiceStats ?? {}),
      ...(localProgress.practiceStats ?? {})
    },
    wordStats: {
      ...(cloudProgress.wordStats ?? {}),
      ...(localProgress.wordStats ?? {})
    },
    reviewQueue: {
      ...(cloudProgress.reviewQueue ?? {}),
      ...(localProgress.reviewQueue ?? {})
    },
    wrongWordProfile: {
      ...(cloudProgress.wrongWordProfile ?? {}),
      ...(localProgress.wrongWordProfile ?? {})
    },
    sightWordProfile: {
      ...(cloudProgress.sightWordProfile ?? {}),
      ...(localProgress.sightWordProfile ?? {})
    },
    grammarSkillProfile: {
      ...(cloudProgress.grammarSkillProfile ?? {}),
      ...(localProgress.grammarSkillProfile ?? {})
    }
  };
}

function mergeCloudStore(localStore, cloudStore) {
  if (!cloudStore || typeof cloudStore !== "object") return localStore;
  const activeWeekId = activeWeekIdForDate();
  const weeklyPlansByWeekId = mergeWeeklyPlansByWeekId(localStore.weeklyPlansByWeekId, cloudStore.weeklyPlansByWeekId);
  const indexedActivePlan = weeklyPlansByWeekId[activeWeekId];
  const weeklyPlan = chooseWeeklyPlanForCloudMerge(
    localStore.weeklyPlan,
    indexedActivePlan || cloudStore.weeklyPlan,
    activeWeekId
  );
  weeklyPlansByWeekId[weeklyPlan.weekId || activeWeekId] = weeklyPlan;
  const merged = {
    ...cloudStore,
    ...localStore,
    weeklyPlan,
    currentWeekId: weeklyPlan.weekId || activeWeekId,
    lastActivatedWeekId: weeklyPlan.weekId || activeWeekId,
    weeklyPlansByWeekId,
    pendingWeeklyPlan: chooseWeeklyPlanForCloudMerge(localStore.pendingWeeklyPlan, cloudStore.pendingWeeklyPlan, nextWeekInfoForDate().weekId),
    profileNames: {
      ...(cloudStore.profileNames ?? {}),
      ...(localStore.profileNames ?? {})
    },
    parentBooks: Array.isArray(localStore.parentBooks) && localStore.parentBooks.length ? localStore.parentBooks : (cloudStore.parentBooks ?? []),
    weeklyArchives: Array.isArray(localStore.weeklyArchives) && localStore.weeklyArchives.length ? localStore.weeklyArchives : (cloudStore.weeklyArchives ?? []),
    wordTestBank: {
      ...(cloudStore.wordTestBank ?? {}),
      ...(localStore.wordTestBank ?? {})
    },
    razArticleDetails: {
      ...(cloudStore.razArticleDetails ?? {}),
      ...(localStore.razArticleDetails ?? {})
    },
    learners: { ...(cloudStore.learners ?? {}), ...(localStore.learners ?? {}) }
  };
  learners.filter((learner) => learner.id !== "tester").forEach((learner) => {
    merged.learners[learner.id] = mergeLearnerProgress(localStore.learners?.[learner.id], cloudStore.learners?.[learner.id]);
  });
  merged.learners.tester = localStore.learners?.tester ?? getDefaultProgress();
  return merged;
}

async function hydrateStoreFromCloud() {
  if (!cloudSyncAvailable() || !window.ReadingBridgeCloud.isSignedIn()) {
    updateCloudSyncStatus("starting", "Cloud sync is still signing in. Try again in a moment.");
    return false;
  }
  try {
    updateCloudSyncStatus("syncing", "Loading cloud progress...");
    const cloudStore = await window.ReadingBridgeCloud.loadCloudStore();
    if (cloudStore) {
      cloudSyncHydrating = true;
      store = mergeCloudStore(store, cloudStore);
      cloudSyncHydrating = false;
      ensureCw25StoryRewardAudio();
      ensureCw26StoryRewardAudio();
      ensureCw26Vocabulary();
      localStorage.setItem(STORAGE_KEY, JSON.stringify(store));
      renderAll();
    }
    await syncStoreToCloud();
    return true;
  } catch (error) {
    cloudSyncHydrating = false;
    updateCloudSyncStatus("sync error", error?.message || "Cloud load failed. Local progress is still available.");
    return false;
  }
}

function setupCloudSync() {
  window.addEventListener("readingbridge:cloud-auth", (event) => {
    const signedIn = Boolean(event.detail?.signedIn);
    if (signedIn) {
      updateCloudSyncStatus("cloud sync on", "Cloud sync on. Hearts and progress will sync automatically.");
      renderParent();
      hydrateStoreFromCloud();
    } else {
      updateCloudSyncStatus("starting", "Starting cloud sync...");
      renderParent();
    }
  });
  window.addEventListener("readingbridge:cloud-error", (event) => {
    updateCloudSyncStatus("sync error", event.detail?.message || "Cloud sync could not start. Local progress is still saved.");
    renderParent();
  });
  if (!cloudSyncAvailable()) {
    updateCloudSyncStatus("loading", "Cloud sync is loading...");
    return;
  }
  window.ReadingBridgeCloud.whenAuthReady().then((user) => {
    if (user) {
      updateCloudSyncStatus("cloud sync on", "Cloud sync on. Hearts and progress will sync automatically.");
      renderParent();
      hydrateStoreFromCloud();
    } else {
      updateCloudSyncStatus("starting", "Starting cloud sync...");
    }
  });
}

function resetTesterForCurrentBuild() {
  const resetKey = `${APP_ASSET_VERSION}:${todayKey()}`;
  if (store.testerDailyResetKey === resetKey) return;
  store.learners.tester = getDefaultProgress();
  store.testerDailyResetKey = resetKey;
  if (state.activeLearnerId === "tester") {
    state.practiceMode = "overview";
    state.activePracticeId = null;
    window.location.hash = "";
  }
  saveStore();
}

function parentIsAuthenticated() {
  return state.parentUnlocked;
}

function profileRequiresPassword(profile) {
  return Boolean(profile?.role === "child" && profile.password);
}

function childIsAuthenticated(profileId) {
  return !profileRequiresPassword(profiles.find((item) => item.id === profileId)) || Boolean(state.childUnlocked[profileId]);
}

function showProfileAuth(profileId, message = "") {
  stopAudio();
  state.authProfileId = profileId;
  state.authMessage = message;
  document.querySelector("#parentPasswordInput").value = "";
  const modal = document.querySelector("#parentAuthModal");
  renderParentAuth();
  modal.hidden = false;
  setTimeout(() => document.querySelector("#parentPasswordInput")?.focus(), 0);
}

function showParentAuth(message = "") {
  showProfileAuth("parent", message);
}

function hideParentAuth() {
  state.authProfileId = "";
  state.authMessage = "";
  document.querySelector("#parentAuthModal").hidden = true;
  document.querySelector("#parentPasswordInput").value = "";
}

function renderParentAuth() {
  const profileId = state.authProfileId || "";
  const profile = profiles.find((item) => item.id === profileId);
  const isChildAuth = profile?.role === "child";
  const form = document.querySelector("#parentAuthForm");
  const eyebrow = document.querySelector("#authEyebrow");
  const intro = document.querySelector("#authIntro");
  const title = document.querySelector("#parentAuthTitle");
  const labelText = document.querySelector("#authPasswordLabelText");
  const input = document.querySelector("#parentPasswordInput");
  const submitButton = document.querySelector("#authSubmitButton");
  form.classList.toggle("child-auth-card", isChildAuth);
  eyebrow.hidden = isChildAuth;
  eyebrow.textContent = isChildAuth ? "" : "Parent access";
  intro.hidden = isChildAuth;
  intro.textContent = isChildAuth ? "" : "Parent pages contain progress, weekly setup, exams, and uploaded learning materials.";
  title.textContent = isChildAuth ? "Unlock" : "Unlock Parent";
  title.classList.toggle("sr-only", isChildAuth);
  labelText.classList.toggle("sr-only", isChildAuth);
  input.placeholder = isChildAuth ? "" : "Enter parent password";
  submitButton.textContent = isChildAuth ? "Unlock" : "Unlock Parent";
  const defaultMessage = "";
  document.querySelector("#parentAuthMessage").textContent = state.authMessage || defaultMessage;
  form.querySelector(".primary-action").disabled = false;
}

async function unlockParent(event) {
  event.preventDefault();
  const profileId = state.authProfileId || "parent";
  const profile = profiles.find((item) => item.id === profileId);
  const password = document.querySelector("#parentPasswordInput").value;
  if (profile?.role === "child") {
    if (password.trim().toLowerCase() !== String(profile.password || "").toLowerCase()) {
      state.authMessage = "Try again.";
      renderParentAuth();
      return;
    }
    state.childUnlocked[profile.id] = true;
    hideParentAuth();
    selectProfile(profile.id);
    return;
  }
  state.authMessage = "Checking password...";
  renderParentAuth();
  const hash = await hashText(password);
  if (hash !== store.parentPasswordHash) {
    state.authMessage = "Wrong password. Please try again.";
    renderParentAuth();
    return;
  }
  state.parentUnlocked = true;
  hideParentAuth();
  selectProfile("parent");
  await loadMaterials();
}

async function logoutParent() {
  state.parentUnlocked = false;
  showProfileGate();
}

async function hashText(text) {
  const data = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function activeProgress() {
  return store.learners[state.activeLearnerId];
}

function activeLearner() {
  const learner = learners.find((item) => item.id === state.activeLearnerId) ?? learners[0];
  return { ...learner, name: profileName(learner.id) };
}

function profileName(profileId) {
  return store.profileNames?.[profileId] ?? profiles.find((profile) => profile.id === profileId)?.defaultName ?? "Learner";
}

function parsePracticeHash() {
  const match = window.location.hash.match(/^#\/kid\/([^/]+)\/practice\/([^/]+)$/);
  if (!match) return null;
  return { learnerId: decodeURIComponent(match[1]), practiceId: decodeURIComponent(match[2]) };
}

function currentPracticeRoute() {
  const route = parsePracticeHash();
  if (!route) return null;
  const learner = learners.find((item) => item.id === route.learnerId);
  const validPracticeIds = new Set(["story", "writing-lab", ...missionCenters.map((item) => item.id), ...adventureLevelIds]);
  return learner && validPracticeIds.has(route.practiceId) ? route : null;
}

function openPracticeRoute(practiceId) {
  window.location.hash = `#/kid/${encodeURIComponent(state.activeLearnerId)}/practice/${encodeURIComponent(practiceId)}`;
}

function parentPracticeRoute(practiceId) {
  if (adventureLevelIds.includes(practiceId)) return "";
  if (["reading", "speak"].includes(practiceId)) return "story";
  if (["listen", "sight", "wordwork"].includes(practiceId)) return "writing-lab";
  return "";
}

function syncPracticeRoute() {
  activateWeeklyPlanForToday();
  ensureCw26StoryRewardAudio();
  ensureCw26Vocabulary();
  const route = currentPracticeRoute();
  if (!route) {
    stopPracticeTimer({ save: true });
    if (state.activeRole === "child") {
      state.practiceMode = "overview";
      state.activePracticeId = null;
      document.body.classList.remove("practice-route-mode", "spelling-route-mode");
      renderAll();
    }
    return;
  }
  stopAudio();
  state.activeLearnerId = route.learnerId;
  state.activeProfileId = route.learnerId;
  state.activeRole = "child";
  state.activeSection = "today";
  state.practiceMode = "route";
  state.activePracticeId = route.practiceId;
  if (missionCenters.some((center) => center.id === route.practiceId) && (!adventureLevelIds.includes(route.practiceId) || isAdventureLevelUnlocked(route.practiceId))) {
    startPracticeTimer(route.practiceId);
  } else {
    stopPracticeTimer({ save: true });
  }
  document.body.classList.add("child-mode");
  document.body.classList.remove("parent-mode");
  document.body.classList.add("practice-route-mode");
  document.body.classList.toggle("spelling-route-mode", route.practiceId === "listen");
  document.querySelector("#profileGate").hidden = true;
  document.querySelector("#siteShell").hidden = false;
  renderAll();
}

function selectProfile(profileId) {
  const profile = profiles.find((item) => item.id === profileId);
  if (!profile) return;
  if (profile.role === "parent" && !parentIsAuthenticated()) {
    showParentAuth();
    return;
  }
  if (profileRequiresPassword(profile) && !childIsAuthenticated(profile.id)) {
    showProfileAuth(profile.id);
    return;
  }
  stopAudio();
  stopPracticeTimer({ save: true });
  activateWeeklyPlanForToday();
  state.activeProfileId = profile.id;
  state.activeRole = profile.role;
  if (profile.role === "child") {
    state.activeLearnerId = profile.id;
    state.activeSection = "today";
    state.practiceMode = "overview";
    state.activePracticeId = null;
    if (currentPracticeRoute()) history.replaceState(null, "", `${window.location.pathname}${window.location.search}`);
  } else {
    state.activeSection = "parent";
  }
  document.body.classList.toggle("child-mode", profile.role === "child");
  document.body.classList.toggle("parent-mode", profile.role === "parent");
  document.body.classList.remove("practice-route-mode", "spelling-route-mode");
  document.querySelector("#profileGate").hidden = true;
  document.querySelector("#siteShell").hidden = false;
  renderAll();
  if (profile.role === "parent") loadMaterials();
}

function showProfileGate() {
  stopAudio();
  stopPracticeTimer({ save: true });
  if (currentPracticeRoute()) history.replaceState(null, "", `${window.location.pathname}${window.location.search}`);
  state.activeProfileId = null;
  state.activeRole = null;
  state.activeSection = "today";
  state.practiceMode = "overview";
  state.activePracticeId = null;
  document.body.classList.remove("child-mode", "parent-mode");
  document.body.classList.remove("practice-route-mode", "spelling-route-mode");
  document.querySelector("#profileGate").hidden = false;
  document.querySelector("#siteShell").hidden = true;
  renderProfileGate();
  window.scrollTo(0, 0);
}

function visibleSections() {
  if (state.activeRole === "parent") return ["parent", "materials", "weekly", "skills", "exam"];
  return ["today", "friday", "review", "practice"];
}

function activePlan() {
  return store.weeklyPlan;
}

function masteredWords() {
  return Object.entries(activeProgress().wordStats)
    .filter(([, stat]) => stat.mastered)
    .map(([word]) => word);
}

function dueReviewWords() {
  const today = todayKey();
  return Object.entries(activeProgress().reviewQueue)
    .filter(([, item]) => item.nextReview <= today)
    .map(([word, item]) => ({ word, ...item }));
}

function practiceKey(centerId, date = todayKey()) {
  return `${date}:${centerId}`;
}

function defaultPracticeStat() {
  return {
    attempts: 0,
    correct: 0,
    completedItems: 0,
    secondsSpent: 0,
    startedAt: null,
    lastUpdated: null
  };
}

function formatDuration(seconds = 0) {
  const total = Math.max(0, Math.floor(seconds));
  const minutes = Math.floor(total / 60);
  const remainder = total % 60;
  return `${minutes}:${String(remainder).padStart(2, "0")}`;
}

function practiceStat(centerId) {
  return { ...defaultPracticeStat(), ...(activeProgress().practiceStats?.[practiceKey(centerId)] ?? {}) };
}

function timedPracticeIsActive(centerId) {
  return state.activePracticeTimerCenterId === centerId && Number.isFinite(state.activePracticeStartedAt);
}

function elapsedPracticeSeconds(centerId) {
  const stat = practiceStat(centerId);
  if (!timedPracticeIsActive(centerId)) return stat.secondsSpent || 0;
  return (stat.secondsSpent || 0) + Math.floor((Date.now() - state.activePracticeStartedAt) / 1000);
}

function updatePracticeTimerDisplay() {
  const centerId = state.activePracticeTimerCenterId;
  const chip = document.querySelector("#practiceTimerChip");
  if (centerId && chip) chip.textContent = formatDuration(elapsedPracticeSeconds(centerId));
}

function startPracticeTimer(centerId) {
  const center = missionCenters.find((item) => item.id === centerId);
  if (center && centerProgress(center).complete) {
    updatePracticeTimerDisplay();
    return;
  }
  if (timedPracticeIsActive(centerId)) {
    updatePracticeTimerDisplay();
    return;
  }
  stopPracticeTimer({ save: true });
  const progress = activeProgress();
  if (!progress.practiceStats) progress.practiceStats = {};
  const key = practiceKey(centerId);
  const now = new Date().toISOString();
  progress.practiceStats[key] = {
    ...defaultPracticeStat(),
    ...(progress.practiceStats[key] ?? {}),
    startedAt: progress.practiceStats[key]?.startedAt || now,
    lastUpdated: now
  };
  state.activePracticeTimerCenterId = centerId;
  state.activePracticeStartedAt = Date.now();
  window.clearInterval(practiceTimerId);
  practiceTimerId = window.setInterval(updatePracticeTimerDisplay, 1000);
  saveStore();
  updatePracticeTimerDisplay();
}

function stopPracticeTimer({ save = true } = {}) {
  if (!state.activePracticeTimerCenterId || !Number.isFinite(state.activePracticeStartedAt)) {
    window.clearInterval(practiceTimerId);
    practiceTimerId = null;
    return;
  }
  const centerId = state.activePracticeTimerCenterId;
  const progress = activeProgress();
  if (!progress.practiceStats) progress.practiceStats = {};
  const key = practiceKey(centerId);
  const current = { ...defaultPracticeStat(), ...(progress.practiceStats[key] ?? {}) };
  const elapsed = Math.max(0, Math.floor((Date.now() - state.activePracticeStartedAt) / 1000));
  progress.practiceStats[key] = {
    ...current,
    secondsSpent: current.secondsSpent + elapsed,
    lastUpdated: new Date().toISOString()
  };
  state.activePracticeTimerCenterId = null;
  state.activePracticeStartedAt = null;
  window.clearInterval(practiceTimerId);
  practiceTimerId = null;
  if (save) saveStore();
}

function weeklyNewWords() {
  return planWeeklyNewWords(activePlan());
}

function defaultAdventureDayProgress(date = todayKey()) {
  return {
    date,
    completedLevels: [],
    readingWarmupRecordingDone: false,
    readingWarmupReplayDone: false,
    readingWarmupSkipped: false,
    razReadingAssignments: [],
    razReadingArticles: {},
    razRetell: defaultRazRetellProgress(),
    wordPowerCompletedWords: [],
    storyRewardComplete: false
  };
}

function ensureAdventureDayProgress(progress = activeProgress(), date = todayKey()) {
  if (!progress.adventureProgress || typeof progress.adventureProgress !== "object") {
    progress.adventureProgress = {};
  }
  const existing = progress.adventureProgress[date];
  const current = existing && typeof existing === "object" ? existing : {};
  Object.assign(current, {
    ...defaultAdventureDayProgress(date),
    ...current
  });
  current.completedLevels = uniqueWords((current.completedLevels ?? []).filter((id) => adventureLevelIds.includes(id)));
  current.wordPowerCompletedWords = uniqueWords((current.wordPowerCompletedWords ?? []).map(displayWord));
  current.razReadingAssignments = uniqueWords(current.razReadingAssignments ?? []);
  current.razReadingArticles = current.razReadingArticles && typeof current.razReadingArticles === "object" ? current.razReadingArticles : {};
  current.razRetell = { ...defaultRazRetellProgress(), ...(current.razRetell ?? {}) };
  if (isSummerRazDate(date)) {
    if (!current.razReadingAssignments.length) {
      current.razReadingAssignments = summerRazAssignmentsForDate(date).map((article) => article.id);
    }
    current.razReadingAssignments.forEach((articleId) => {
      current.razReadingArticles[articleId] = {
        recordingDone: false,
        replayDone: false,
        ...(current.razReadingArticles[articleId] ?? {})
      };
    });
    if (current.razReadingAssignments.length === 1 && !current.razRetell.selectedArticleId) {
      current.razRetell.selectedArticleId = current.razReadingAssignments[0];
      current.razRetell.selectedAt = new Date().toISOString();
    }
    const articleStates = current.razReadingAssignments.map((articleId) => current.razReadingArticles[articleId]);
    current.readingWarmupRecordingDone = articleStates.length > 0 && articleStates.every((article) => article.recordingDone);
    current.readingWarmupReplayDone = articleStates.length > 0 && articleStates.every((article) => article.replayDone);
  }
  progress.adventureProgress[date] = current;
  syncLegacyAdventureCompletions(progress, date);
  syncWordPowerAdventureCompletion(progress, date);
  return progress.adventureProgress[date];
}

function syncWordPowerAdventureCompletion(progress = activeProgress(), date = todayKey()) {
  const dayProgress = progress.adventureProgress?.[date];
  if (!dayProgress || dayProgress.completedLevels.includes("word-power")) return;
  const words = weeklyVocabulary();
  const target = wordPowerTarget(words);
  if (!target) return;
  const completedKeys = new Set((dayProgress.wordPowerCompletedWords ?? []).map(normalizeWord));
  const availableKeys = new Set(words.map((item) => normalizeWord(item.word)));
  const count = [...completedKeys].filter((word) => availableKeys.has(word)).length;
  if (count >= target) {
    dayProgress.completedLevels.push("word-power");
  }
}

function syncLegacyAdventureCompletions(progress = activeProgress(), date = todayKey()) {
  const dayProgress = progress.adventureProgress?.[date];
  if (!dayProgress) return;
  Object.entries(legacyCenterAdventureMap).forEach(([centerId, levelId]) => {
    const center = missionCenters.find((item) => item.id === centerId);
    if (!center) return;
    const stat = { ...defaultPracticeStat(), ...(progress.practiceStats?.[practiceKey(centerId, date)] ?? {}) };
    const spellingSession = centerId === "listen" ? progress.spellingSessions?.[date] : null;
    const goal = centerId === "listen"
      ? (spellingSession?.words?.length || dailySpellingTarget(progress, new Date(`${date}T00:00:00`)))
      : centerId === "sight"
        ? (progress.sightSessions?.[date]?.activities?.length || progress.sightSessions?.[date]?.words?.length || sightPracticeWords(progress).length)
        : center.completionGoal || 1;
    const complete = centerId === "listen"
      ? spellingSessionIsComplete(spellingSession) || (stat.completedItems ?? 0) >= goal
      : (stat.completedItems ?? 0) >= goal;
    if (goal && complete && !dayProgress.completedLevels.includes(levelId)) {
      dayProgress.completedLevels.push(levelId);
    }
  });
  if (dayProgress.storyRewardComplete && !dayProgress.completedLevels.includes("story-reward")) {
    dayProgress.completedLevels.push("story-reward");
  }
}

function completeAdventureLevel(levelId, { render = true } = {}) {
  if (!adventureLevelIds.includes(levelId)) return { status: "unknown", completed: false, unlocked: false };
  const progress = activeProgress();
  const dayProgress = ensureAdventureDayProgress(progress);
  if (!isAdventureLevelUnlocked(levelId)) {
    return { status: "locked", completed: false, unlocked: false };
  }
  if (levelId === "read-warmup" && !dayProgress.readingWarmupSkipped && (!dayProgress.readingWarmupRecordingDone || !dayProgress.readingWarmupReplayDone)) {
    return {
      status: "next",
      completed: false,
      unlocked: true,
      readingWarmupRecordingDone: Boolean(dayProgress.readingWarmupRecordingDone),
      readingWarmupReplayDone: Boolean(dayProgress.readingWarmupReplayDone)
    };
  }
  if (levelId === "word-power" && !getWordPowerStatus().ready) {
    return { status: "next", completed: false, unlocked: true, ...getWordPowerStatus() };
  }
  if (!dayProgress.completedLevels.includes(levelId)) {
    dayProgress.completedLevels.push(levelId);
  }
  if (levelId === "story-reward") dayProgress.storyRewardComplete = true;
  saveStore();
  if (render) renderAll();
  return getTodayAdventureLevels().find((level) => level.id === levelId);
}

function isAdventureLevelCompleted(levelId, date = todayKey(), progress = activeProgress()) {
  return ensureAdventureDayProgress(progress, date).completedLevels.includes(levelId);
}

function isAdventureLevelUnlocked(levelId, date = todayKey(), progress = activeProgress()) {
  const index = adventureLevelDefinitions.findIndex((level) => level.id === levelId);
  if (index < 0) return false;
  if (index === 0) return true;
  const dayProgress = ensureAdventureDayProgress(progress, date);
  const previousLevelId = adventureLevelDefinitions[index - 1].id;
  return dayProgress.completedLevels.includes(previousLevelId);
}

function getTodayAdventureLevels() {
  ensureAdventureDayProgress(activeProgress());
  const firstIncompleteIndex = adventureLevelDefinitions.findIndex((level) => !isAdventureLevelCompleted(level.id));
  return adventureLevelDefinitions.map((level, index) => {
    const completed = isAdventureLevelCompleted(level.id);
    const unlocked = isAdventureLevelUnlocked(level.id);
    return {
      ...level,
      status: completed ? "done" : unlocked && (firstIncompleteIndex < 0 || index === firstIncompleteIndex) ? "next" : "locked",
      completed,
      unlocked
    };
  });
}

function wordPowerTarget(words = weeklyVocabulary()) {
  return words.length;
}

function recordWordPowerWord(word, { render = true } = {}) {
  const key = normalizeWord(word);
  if (!key) return getWordPowerStatus();
  const dayProgress = ensureAdventureDayProgress(activeProgress());
  const existingKeys = new Set((dayProgress.wordPowerCompletedWords ?? []).map(normalizeWord));
  if (!existingKeys.has(key)) {
    dayProgress.wordPowerCompletedWords = [...(dayProgress.wordPowerCompletedWords ?? []), displayWord(word)];
  }
  saveStore();
  if (render) renderAll();
  return getWordPowerStatus();
}

function getWordPowerStatus() {
  const words = weeklyVocabulary();
  const dayProgress = ensureAdventureDayProgress(activeProgress());
  const completedKeys = new Set((dayProgress.wordPowerCompletedWords ?? []).map(normalizeWord));
  const availableKeys = new Set(words.map((item) => normalizeWord(item.word)));
  const count = [...completedKeys].filter((word) => availableKeys.has(word)).length;
  const target = wordPowerTarget(words);
  return {
    completedWords: [...completedKeys],
    count,
    target,
    ready: target > 0 && count >= target,
    complete: isAdventureLevelCompleted("word-power")
  };
}

function completeWordPower({ render = true } = {}) {
  const status = getWordPowerStatus();
  if (!status.ready) return { status: "next", completed: false, unlocked: isAdventureLevelUnlocked("word-power"), ...status };
  const dayProgress = ensureAdventureDayProgress(activeProgress());
  if (!dayProgress.completedLevels.includes("word-power")) {
    dayProgress.completedLevels.push("word-power");
  }
  saveStore();
  const completedLevel = getTodayAdventureLevels().find((level) => level.id === "word-power");
  const result = {
    ...status,
    ...completedLevel,
    completed: Boolean(completedLevel?.completed),
    complete: Boolean(completedLevel?.completed)
  };
  if (render) renderAll();
  return result;
}

function completeReadingWarmup({ render = true } = {}) {
  const dayProgress = ensureAdventureDayProgress(activeProgress());
  if (dayProgress.readingWarmupSkipped) {
    stopReadingWarmupReplay({ render: false });
    return completeAdventureLevel("read-warmup", { render });
  }
  if (!dayProgress.readingWarmupRecordingDone || !dayProgress.readingWarmupReplayDone) {
    return {
      status: "next",
      completed: false,
      unlocked: isAdventureLevelUnlocked("read-warmup"),
      readingWarmupRecordingDone: Boolean(dayProgress.readingWarmupRecordingDone),
      readingWarmupReplayDone: Boolean(dayProgress.readingWarmupReplayDone)
    };
  }
  stopReadingWarmupReplay({ render: false });
  return completeAdventureLevel("read-warmup", { render });
}

function readingWarmupIsRequired(date = todayKey()) {
  const day = new Date(`${date}T00:00:00`).getDay();
  return day === 0 || day === 1;
}

function readingWarmupSkipCount(progress = activeProgress(), date = todayKey()) {
  const weekId = activeWeekIdForDate(new Date(`${date}T00:00:00`));
  return Object.entries(progress.adventureProgress ?? {})
    .filter(([dayKey, dayProgress]) => {
      if (!dayProgress?.readingWarmupSkipped) return false;
      return activeWeekIdForDate(new Date(`${dayKey}T00:00:00`)) === weekId;
    })
    .length;
}

function readingWarmupSkipStatus(progress = activeProgress(), date = todayKey()) {
  const required = readingWarmupIsRequired(date);
  const used = readingWarmupSkipCount(progress, date);
  const limit = 3;
  return {
    required,
    used,
    limit,
    remaining: Math.max(0, limit - used),
    canSkip: !required && used < limit
  };
}

function skipReadingWarmup({ render = true } = {}) {
  const progress = activeProgress();
  const date = todayKey();
  const status = readingWarmupSkipStatus(progress, date);
  const dayProgress = ensureAdventureDayProgress(progress, date);
  if (!status.canSkip && !dayProgress.readingWarmupSkipped) {
    return { status: "required", completed: false, unlocked: true, ...status };
  }
  dayProgress.readingWarmupSkipped = true;
  if (!dayProgress.completedLevels.includes("read-warmup")) {
    dayProgress.completedLevels.push("read-warmup");
  }
  saveStore();
  if (render) renderAll();
  return getTodayAdventureLevels().find((level) => level.id === "read-warmup");
}

function completeStoryReward({ render = true } = {}) {
  const result = completeAdventureLevel("story-reward", { render: false });
  const heartEarned = result?.completed ? awardDailyHeart() : false;
  if (heartEarned) {
    state.heartRewardOverlay = { count: heartCount() };
  }
  if (render) renderAll();
  return { ...result, heartEarned };
}

function returnToAdventureMap() {
  stopAudio();
  stopReadingWarmupReplay({ render: false });
  state.practiceMode = "overview";
  state.activePracticeId = null;
  window.location.hash = "";
  renderAll();
}

function heartCount(progress = activeProgress()) {
  return Array.isArray(progress.heartRewardDays) ? uniqueWords(progress.heartRewardDays).length : 0;
}

function awardDailyHeart(progress = activeProgress(), date = todayKey()) {
  if (!Array.isArray(progress.heartRewardDays)) progress.heartRewardDays = [];
  if (progress.heartRewardDays.includes(date)) return false;
  progress.heartRewardDays.push(date);
  saveStore();
  return true;
}

function isSundaySpellingReview(date = new Date()) {
  return date.getDay() === 0;
}

function wrongProfileWords(progress = activeProgress()) {
  return Object.entries(progress.wrongWordProfile ?? {})
    .filter(([, item]) => item.status !== "mastered")
    .sort((a, b) => (b[1].missCount ?? 0) - (a[1].missCount ?? 0))
    .map(([word, item]) => displayWord(item.word || word));
}

function sundaySpellingReviewWords(progress = activeProgress()) {
  const ranked = new Map();
  const addWord = (word, score = 1) => {
    const key = normalizeWord(word);
    if (!key) return;
    const current = ranked.get(key) ?? { word: displayWord(word), score: 0 };
    ranked.set(key, {
      word: current.word || displayWord(word),
      score: current.score + score
    });
  };

  Object.entries(progress.wrongWordProfile ?? {}).forEach(([word, item]) => {
    if (item?.status === "mastered") return;
    addWord(item?.word || word, 100 + (item?.missCount ?? 0));
  });

  Object.values(progress.spellingSessions ?? {}).forEach((session) => {
    (session?.results ?? []).forEach((result) => {
      if (result?.correct === false) addWord(result.word, 60 + (result.attempts ?? 0));
    });
  });

  Object.entries(progress.wordStats ?? {}).forEach(([word, stat]) => {
    if (stat?.mastered) return;
    const attempts = stat?.attempts ?? 0;
    const correct = stat?.correct ?? 0;
    if (stat?.lastResult === "practice" || attempts > correct) addWord(word, 30 + attempts - correct);
  });

  return [...ranked.values()]
    .sort((a, b) => b.score - a.score || a.word.localeCompare(b.word))
    .map((item) => item.word);
}

function spellingReviewWords(progress = activeProgress()) {
  const weeklyKeys = new Set(weeklyNewWords().map(normalizeWord));
  return uniqueWords([
    ...wrongProfileWords(progress),
    ...testBankWords(["spelling", "phonicsReview", "mastery"]).slice(0, 8)
  ]).filter((word) => !weeklyKeys.has(normalizeWord(word)));
}

function spellingPool(progress = activeProgress()) {
  return uniqueWords([...weeklyNewWords(), ...spellingReviewWords(progress)]);
}

function spellingSourceWordsForDate(progress = activeProgress(), date = new Date()) {
  if (isSundaySpellingReview(date)) return sundaySpellingReviewWords(progress);
  return spellingPool(progress);
}

function sightWordBankWords() {
  return uniqueWords(sightWordSeedBank.map((item) => item.word));
}

function weeklySightWords() {
  return uniqueWords([...(activePlan().highFrequencyWords ?? [])]).filter(Boolean);
}

function sightWordEntry(word) {
  const key = normalizeWord(word);
  return sightWordSeedBank.find((item) => normalizeWord(item.word) === key) ?? {
    word: displayWord(word),
    grade: "weekly",
    source: "weekly plan",
    isIrregular: true,
    heartPart: "",
    definition: wordHelp[key]?.[0] ?? "a word good readers know quickly",
    example: wordHelp[key]?.[1] ?? meaningfulSightSentence(word)
  };
}

function sightPracticeWords(progress = activeProgress()) {
  const weekly = weeklySightWords();
  const weak = Object.entries(progress.sightWordProfile ?? {})
    .filter(([, stat]) => (stat.missed ?? 0) > (stat.correct ?? 0))
    .sort((a, b) => (b[1].missed ?? 0) - (a[1].missed ?? 0))
    .map(([word]) => displayWord(word));
  return uniqueWords(weekly.length ? [...weak.slice(0, 2), ...weekly] : weak).slice(0, MAX_DAILY_SIGHT_WORDS);
}

function sightChoicePool(word, session = ensureSightSession()) {
  const key = normalizeWord(word);
  const pool = uniqueWords([
    ...(session.words ?? []),
    ...weeklySightWords(),
    ...sightWordBankWords()
  ]).filter((item) => normalizeWord(item) !== key);
  return pool;
}

function sightActivitiesForWords(words, date = todayKey()) {
  const activityTypes = ["hear-find", "sentence-cloze", "word-build", "mini-passage-hunt", "sentence-match"];
  return words.map((word, index) => {
    const type = activityTypes[seededValue(`${state.activeLearnerId}:${date}:sight-activity:${word}:${index}`) % activityTypes.length];
    return { id: `${normalizeWord(word)}-${index}`, word: displayWord(word), type };
  });
}

function ensureSightSession() {
  const progress = activeProgress();
  const date = todayKey();
  if (!progress.sightSessions) progress.sightSessions = {};
  if (!progress.sightWordProfile) progress.sightWordProfile = {};
  const weekly = weeklySightWords();
  const weak = Object.entries(progress.sightWordProfile)
    .filter(([, stat]) => (stat.missed ?? 0) > (stat.correct ?? 0))
    .sort((a, b) => (b[1].missed ?? 0) - (a[1].missed ?? 0))
    .map(([word]) => displayWord(word));
  const sourceKey = `sight-v2:${activePlan().weekId}:${date}:high-frequency:${weekly.map(normalizeWord).join("|")}`;
  const saved = progress.sightSessions[date];
  if (saved?.sourceKey === sourceKey) return saved;
  const reviewWords = weak.slice(0, weekly.length ? 2 : MAX_DAILY_SIGHT_WORDS);
  const mainWords = weekly.length ? seededShuffle(weekly, `${state.activeLearnerId}:${date}:weekly-sight`) : [];
  const words = uniqueWords([
    ...reviewWords,
    ...mainWords
  ]).slice(0, MAX_DAILY_SIGHT_WORDS);
  const activities = sightActivitiesForWords(words, date);
  progress.sightSessions[date] = {
    date,
    sourceKey,
    words,
    activities,
    currentIndex: 0,
    completedWords: [],
    completedActivities: [],
    results: [],
    modeIndex: 0,
    buildState: null,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  saveStore();
  return progress.sightSessions[date];
}

function dailySpellingTarget(progress = activeProgress(), date = new Date()) {
  const poolCount = spellingSourceWordsForDate(progress, date).length;
  if (!poolCount) return 0;
  return Math.min(
    MAX_DAILY_SPELLING_WORDS,
    poolCount,
    Math.max(MIN_DAILY_SPELLING_WORDS, Math.ceil(poolCount * DAILY_SPELLING_SOURCE_RATIO))
  );
}

function spellingDaysRemainingThroughThursday(date = new Date()) {
  const day = date.getDay();
  if (day === 0) return 5;
  if (day >= 5) return 1;
  return Math.max(1, 5 - day);
}

function requiredNewWordsForThursdayCoverage(weeklyWords, exposure, date = new Date()) {
  const remainingExposures = weeklyWords.reduce((total, word) => {
    const count = exposure[normalizeWord(word)] ?? 0;
    return total + Math.max(0, WEEKLY_NEW_WORD_EXPOSURE_GOAL - count);
  }, 0);
  return Math.ceil(remainingExposures / spellingDaysRemainingThroughThursday(date));
}

function seededValue(seed) {
  let value = 2166136261;
  for (let index = 0; index < seed.length; index += 1) {
    value ^= seed.charCodeAt(index);
    value = Math.imul(value, 16777619);
  }
  return value >>> 0;
}

function seededShuffle(items, seed) {
  const list = [...items];
  let value = seededValue(seed);
  for (let index = list.length - 1; index > 0; index -= 1) {
    value = (value * 1664525 + 1013904223) >>> 0;
    const swapIndex = value % (index + 1);
    [list[index], list[swapIndex]] = [list[swapIndex], list[index]];
  }
  return list;
}

function weeklyVideoSearchUrl(plan = activePlan()) {
  const title = plan.readingArticleTitle || plan.readingTitle || "read aloud for kids";
  return `https://www.youtube.com/results?search_query=${encodeURIComponent(`${title} read aloud for kids`)}`;
}

function schoolReadingWarmupSource(plan = activePlan()) {
  const sequenceSelection = Array.isArray(plan.weeklySequence?.selections) ? plan.weeklySequence.selections[0] : "";
  const title = plan.readingArticleTitle || plan.readingTitle || sequenceSelection || "this week's school reading";
  return {
    title,
    summary: plan.readingArticleSummary || `Read ${title} from your school book.`,
    sourceLabel: sequenceSelection && normalizeWord(sequenceSelection) !== normalizeWord(title) ? sequenceSelection : "",
    videoUrl: weeklyVideoSearchUrl({ ...plan, readingArticleTitle: title, readingTitle: title })
  };
}

function storyKeywords(plan = activePlan()) {
  const vocabularyWords = Array.isArray(plan.weeklyVocabulary) ? plan.weeklyVocabulary.map((item) => item.word) : [];
  const storyFriendlyWords = vocabularyWords.filter((word) => {
    const key = normalizeWord(word);
    return key.length > 2 && !key.includes("'") && !["conducting", "spoiled", "rummaged", "merrily"].includes(key);
  });
  return uniqueWords([
    ...storyFriendlyWords,
    ...(plan.masteryWords ?? []).filter((word) => normalizeWord(word).length > 2 && !normalizeWord(word).includes("'"))
  ]).filter(Boolean);
}

function makeChoiceQuestion(prompt, answer, distractors, seed) {
  const choices = seededShuffle(uniqueWords([answer, ...distractors]).slice(0, 4), seed);
  return {
    prompt,
    choices,
    answer,
    answerIndex: choices.findIndex((choice) => choice === answer)
  };
}

function storyWordExposureKey(plan = activePlan()) {
  return plan.weekId || weekKey();
}

function storyExposureCount(word, progress = activeProgress(), plan = activePlan()) {
  const key = storyWordExposureKey(plan);
  return progress.storyWordExposure?.[key]?.[normalizeWord(word)] ?? 0;
}

function selectStoryWords(plan = activePlan(), date = todayKey(), learnerId = state.activeLearnerId, progress = activeProgress()) {
  const candidates = storyKeywords(plan).filter((word) => normalizeWord(word).length > 1);
  const underTarget = candidates.filter((word) => storyExposureCount(word, progress, plan) < WEEKLY_NEW_WORD_EXPOSURE_GOAL);
  const pool = underTarget.length ? underTarget : candidates;
  const prioritized = pool.sort((a, b) => storyExposureCount(a, progress, plan) - storyExposureCount(b, progress, plan));
  return seededShuffle(prioritized.slice(0, 16), `${learnerId}:${date}:story-words`).slice(0, 3);
}

function storyWordPhrase(words) {
  const list = words.filter(Boolean).slice(0, 3);
  return {
    first: list[0] || "look",
    second: list[1] || "good",
    third: list[2] || "safe"
  };
}

function svgDataUri(markup) {
  return `data:image/svg+xml;charset=UTF-8,${encodeURIComponent(markup)}`;
}

function storyCover(storyTitle) {
  const covers = {
    "The Quiet Blue": {
      alt: "A torn blue paper becomes a river under a small bird.",
      svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 160"><rect width="240" height="160" rx="20" fill="#dff6ff"/><path d="M0 116 C42 86 72 142 118 103 C153 73 181 104 240 76 L240 160 L0 160 Z" fill="#9fd7ef"/><path d="M119 0 L104 38 L125 62 L110 96 L134 132 L121 160" fill="none" stroke="#ffffff" stroke-width="8" stroke-linecap="round"/><path d="M42 54 q12 -14 24 0 q12 -14 24 0" fill="none" stroke="#365c6b" stroke-width="5" stroke-linecap="round"/></svg>`
    },
    "The Day the Wall Listened": {
      alt: "A classroom wall with a painted path and a bright window.",
      svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 160"><rect width="240" height="160" rx="20" fill="#fff5d8"/><rect x="28" y="26" width="184" height="86" rx="14" fill="#f7fbff" stroke="#6fb6c8" stroke-width="6"/><rect x="48" y="42" width="38" height="30" rx="6" fill="#ffd36b"/><path d="M42 132 C82 104 114 148 152 116 C172 99 196 104 220 92" fill="none" stroke="#ef8a6b" stroke-width="12" stroke-linecap="round"/><circle cx="154" cy="58" r="10" fill="#7aa37a"/></svg>`
    },
    "A Pocket for the Wind": {
      alt: "A small paper kite lifting in a gentle wind.",
      svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 160"><rect width="240" height="160" rx="20" fill="#eaf8ef"/><path d="M62 52 L118 24 L172 54 L118 98 Z" fill="#ffd36b" stroke="#315a7d" stroke-width="5"/><path d="M118 98 C126 120 94 122 108 142" fill="none" stroke="#315a7d" stroke-width="4"/><path d="M32 46 C62 22 82 70 112 46 M142 34 C170 16 184 54 214 36" fill="none" stroke="#8edbe7" stroke-width="7" stroke-linecap="round"/><circle cx="58" cy="116" r="10" fill="#ef8a6b"/></svg>`
    },
    "The Smallest Color": {
      alt: "A tiny yellow dot grows into a rainy town picture.",
      svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 160"><rect width="240" height="160" rx="20" fill="#eef4fb"/><path d="M36 28 v34 M74 18 v44 M198 26 v36" stroke="#8edbe7" stroke-width="6" stroke-linecap="round"/><circle cx="120" cy="82" r="17" fill="#ffd36b"/><path d="M54 134 h132 l-20 -34 h-92 Z" fill="#7aa37a"/><rect x="92" y="103" width="20" height="31" fill="#fff"/><rect x="130" y="103" width="20" height="31" fill="#fff"/></svg>`
    },
    "When the Box Hummed": {
      alt: "A cardboard box opens into a tiny moonlit theater.",
      svg: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 160"><rect width="240" height="160" rx="20" fill="#fff2e8"/><path d="M54 58 h132 v76 H54 Z" fill="#d9a15f" stroke="#8a623c" stroke-width="6"/><path d="M54 58 L84 30 h102 l-30 28 Z" fill="#e8bd7b" stroke="#8a623c" stroke-width="6"/><circle cx="146" cy="86" r="20" fill="#fff7c2"/><path d="M82 112 C106 92 132 130 160 106" fill="none" stroke="#315a7d" stroke-width="6" stroke-linecap="round"/></svg>`
    }
  };
  const cover = covers[storyTitle] ?? covers["The Quiet Blue"];
  return { alt: cover.alt, imageDataUri: svgDataUri(cover.svg) };
}

function networkStoryQuestion(prompt, answer, choices) {
  const uniqueChoices = uniqueWords([answer, ...choices]).slice(0, 4);
  return {
    prompt,
    choices: uniqueChoices,
    answer,
    answerIndex: uniqueChoices.findIndex((choice) => choice === answer)
  };
}

function normalizeListeningStory(story) {
  if (!story || typeof story !== "object" || !story.title || !story.url) return null;
  const questions = Array.isArray(story.questions) ? story.questions.filter((question) => question?.prompt && question?.answer) : [];
  const normalizedQuestions = questions.map((question) => {
    const choices = Array.isArray(question.choices) && question.choices.length ? question.choices : [question.answer];
    const answerIndex = choices.findIndex((choice) => choice === question.answer);
    return {
      prompt: question.prompt,
      choices,
      answer: question.answer,
      answerIndex: answerIndex >= 0 ? answerIndex : 0
    };
  });
  const cover = story.cover ?? storyCover(story.title);
  return {
    dayIndex: Number.isFinite(Number(story.dayIndex)) ? Number(story.dayIndex) : 0,
    storyTitle: story.storyTitle || story.title,
    title: story.title,
    sourceName: story.sourceName || "Online story",
    url: story.url,
    audioUrl: story.audioUrl || "",
    durationSeconds: Number.isFinite(Number(story.durationSeconds)) ? Number(story.durationSeconds) : null,
    coverUrl: story.coverUrl || "",
    cover,
    summary: story.summary || "",
    storyText: story.storyText || story.summary || `Open ${story.title} to listen to the story.`,
    skillMatch: story.skillMatch || "",
    questions: normalizedQuestions,
    usedWords: Array.isArray(story.usedWords) ? story.usedWords : [],
    keywords: Array.isArray(story.keywords) ? story.keywords : []
  };
}

function hasPlayableAudio(story) {
  if (!story?.audioUrl) return false;
  if (Number.isFinite(Number(story.durationSeconds)) && Number(story.durationSeconds) > STORY_REWARD_MAX_AUDIO_SECONDS) return false;
  return /^(https?:\/\/|assets\/|\.\/|\/).+\.(mp3|m4a|mp4|wav)(\?.*)?$/i.test(story.audioUrl);
}

function cw26ListeningStories() {
  const question = (prompt, answer, choices) => networkStoryQuestion(prompt, answer, choices);
  const story = ({ dayIndex, title, sourceName, url, audioUrl, durationSeconds, summary, skillMatch, usedWords, questions }) => normalizeListeningStory({
    dayIndex,
    title,
    storyTitle: title,
    sourceName,
    url,
    audioUrl,
    durationSeconds,
    summary,
    storyText: summary,
    skillMatch,
    usedWords,
    keywords: uniqueWords([title, ...usedWords]),
    questions
  });
  return [
    story({
      dayIndex: 0,
      title: "When Your Trainers Go for a Run",
      sourceName: "Storynory",
      url: "https://podcasts.apple.com/us/podcast/when-your-trainers-go-for-a-run/id94571049?i=1000758377636&uo=4",
      audioUrl: "https://traffic.libsyn.com/secure/blogrelations/hugh-trainers-storynory-2-16.mp3?dest-id=10605",
      durationSeconds: 240.48,
      summary: "A playful fantasy about a pair of trainers that go running on their own while their owner wonders where they are.",
      skillMatch: "Fantasy, character actions, problem and solution.",
      usedWords: ["trainers", "running", "missing", "fantasy"],
      questions: [
        question("What goes for a run in the story?", "a pair of trainers", ["a school bus", "a red kite", "a toy boat"]),
        question("Why are the trainers hard to find?", "they went running on their own", ["they are under the bed", "a dog buried them", "they fell in a box"]),
        question("What are trainers?", "shoes", ["hats", "gloves", "coats"]),
        question("What kind of story is this?", "a fantasy", ["a weather report", "a recipe", "a true news story"])
      ]
    }),
    story({
      dayIndex: 1,
      title: "The Fox Princess",
      sourceName: "Storynory",
      url: "https://podcasts.apple.com/us/podcast/the-fox-princess/id94571049?i=1000750003709&uo=4",
      audioUrl: "https://traffic.libsyn.com/secure/blogrelations/fox-princess-16-2.mp3?dest-id=10605",
      durationSeconds: 271.52,
      summary: "A shape-shifting fox becomes a princess, explores a palace, plays hide-and-seek, and races away when the king's dogs appear.",
      skillMatch: "Character, setting, sequence, and fantasy details.",
      usedWords: ["fox", "princess", "palace", "garden"],
      questions: [
        question("Why does the fox turn into a princess?", "she does not want a guard to spot her", ["she wants to fly", "she loses her tail", "she wants to paint"]),
        question("What food does the fox princess ask for?", "chicken wings and warm milk", ["bread and soup", "cake and tea", "apples and water"]),
        question("What game does she play with the princeling?", "hide-and-seek", ["football", "chess", "tag"]),
        question("Why does the fox run out of the palace?", "the king's dogs come into the garden", ["the palace closes", "it begins to snow", "the cook calls her"])
      ]
    }),
    story({
      dayIndex: 2,
      title: "The Camel Who Learned to Trust the Journey",
      sourceName: "Little Dreamers",
      url: "https://podcasts.apple.com/us/podcast/the-camel-who-learned-to-trust-the-journey-a/id1836989412?i=1000728479316&uo=4",
      audioUrl: "https://traffic.libsyn.com/secure/f383a122-7b84-43ba-a074-80dc49d52393/Little_Dreamers_-_ep_9.mp3?dest-id=5139900",
      durationSeconds: 258.06,
      summary: "Cassia the Camel worries about crossing the desert until a wise fox and her caravan help her find courage one step at a time.",
      skillMatch: "Character feelings, advice, courage, and lesson.",
      usedWords: ["camel", "desert", "courage", "journey"],
      questions: [
        question("What is Cassia worried about crossing?", "the desert", ["a river", "a city", "a forest"]),
        question("Who teaches Cassia about courage?", "a wise desert fox", ["a noisy parrot", "a young rabbit", "a sleepy bear"]),
        question("Who supports Cassia on the journey?", "her caravan", ["a band", "a football team", "a class"]),
        question("What does Cassia learn?", "courage can grow one step at a time", ["the whole path must be known", "brave animals never worry", "traveling alone is best"])
      ]
    }),
    story({
      dayIndex: 3,
      title: "The Fox Who Shared Her Fireflies",
      sourceName: "Little Dreamers",
      url: "https://podcasts.apple.com/us/podcast/the-fox-who-shared-her-fireflies/id1836989412?i=1000725280160&uo=4",
      audioUrl: "https://traffic.libsyn.com/secure/f383a122-7b84-43ba-a074-80dc49d52393/Little_Dreamers_-_ep_4.mp3?dest-id=5139900",
      durationSeconds: 248.24,
      summary: "Fern the Fox keeps a lantern of fireflies to herself until a dark night teaches her that sharing can help everyone.",
      skillMatch: "Problem and solution, generosity, and central message.",
      usedWords: ["fox", "fireflies", "lantern", "sharing"],
      questions: [
        question("What does Fern keep in her secret lantern?", "fireflies", ["snowflakes", "seashells", "buttons"]),
        question("When do Fern's friends need her light?", "on a dark night", ["at lunchtime", "on a sunny morning", "during a picnic"]),
        question("Who helps Fern understand sharing?", "a wise badger", ["a little fish", "a green frog", "a tall giraffe"]),
        question("What lesson does Fern learn?", "sharing makes the world brighter", ["secrets solve every problem", "only foxes need light", "friends should work alone"])
      ]
    }),
    story({
      dayIndex: 4,
      title: "The Lion Who Learned to Listen",
      sourceName: "Little Dreamers",
      url: "https://podcasts.apple.com/us/podcast/the-lion-who-learned-to-listen/id1836989412?i=1000724992943&uo=4",
      audioUrl: "https://traffic.libsyn.com/secure/f383a122-7b84-43ba-a074-80dc49d52393/Little_Dreamers_-_ep_3.mp3?dest-id=5139900",
      durationSeconds: 266.06,
      summary: "Leo the Lion roars so much that he misses his friends' wisdom and warnings, then learns that listening is a kind of strength.",
      skillMatch: "Cause and effect, character change, and listening.",
      usedWords: ["lion", "listen", "warning", "strength"],
      questions: [
        question("What does Leo do too much?", "he roars", ["he whispers", "he sleeps", "he swims"]),
        question("What does Leo miss because he is noisy?", "his friends' wisdom and warnings", ["his breakfast", "a school bus", "a birthday cake"]),
        question("Which animals help Leo learn?", "a tortoise and meerkats", ["a whale and dolphins", "cats and dogs", "penguins and seals"]),
        question("What does Leo learn about strength?", "strong leaders listen", ["loud voices are always right", "friends should stay quiet", "warnings do not matter"])
      ]
    }),
    story({
      dayIndex: 5,
      title: "The Singing Bird",
      sourceName: "Brainytoon",
      url: "https://podcasts.apple.com/us/podcast/the-singing-bird/id1444798158?i=1000613621119&uo=4",
      audioUrl: "https://traffic.libsyn.com/secure/brainytoon/4-POD-ENG-Singing-Bird.mp3?dest-id=876437",
      durationSeconds: 253.02,
      summary: "A singing bird faces a challenge and learns not to stop doing what she loves.",
      skillMatch: "Character motivation, perseverance, and lesson.",
      usedWords: ["bird", "singing", "challenge", "perseverance"],
      questions: [
        question("What does the bird love to do?", "sing", ["paint", "swim", "build"]),
        question("What kind of animal is the main character?", "a bird", ["a rabbit", "a turtle", "a horse"]),
        question("What does the bird decide not to do?", "give up singing", ["leave her nest", "talk to friends", "look for food"]),
        question("What is the story's lesson?", "keep doing what you love", ["stop when work is hard", "hide your talents", "never ask for help"])
      ]
    }),
    story({
      dayIndex: 6,
      title: "The Sparrow Who Learned Not to Worry",
      sourceName: "Little Dreamers",
      url: "https://podcasts.apple.com/us/podcast/the-sparrow-who-learned-not-to-worry-a-bedtime/id1836989412?i=1000729591940&uo=4",
      audioUrl: "https://traffic.libsyn.com/secure/f383a122-7b84-43ba-a074-80dc49d52393/Little_Dreamers_-_ep_10.mp3?dest-id=5139900",
      durationSeconds: 297.27,
      summary: "Sera the Sparrow worries about storms, food, and tomorrow until a wise owl helps her let go of what-ifs and enjoy today.",
      skillMatch: "Character feelings, advice, mindfulness, and lesson.",
      usedWords: ["sparrow", "worry", "owl", "today"],
      questions: [
        question("What kind of bird is Sera?", "a sparrow", ["an eagle", "a penguin", "a swan"]),
        question("What does Sera worry about?", "storms, food, and tomorrow", ["games and songs", "paint and paper", "boats and trains"]),
        question("Who helps Sera?", "a wise owl", ["a young mouse", "a busy bee", "a blue whale"]),
        question("What does Sera learn to enjoy?", "the day she is in", ["every possible what-if", "worrying about tomorrow", "carrying every fear"])
      ]
    })
  ].filter(hasPlayableAudio);
}

function cw25Question(prompt, answer, choices) {
  return networkStoryQuestion(prompt, answer, choices);
}

function cw25ListeningStories() {
  const story = (dayIndex, title, summary, skillMatch, storyText, usedWords, questions, durationSeconds) => normalizeListeningStory({
    dayIndex,
    title,
    storyTitle: title,
    sourceName: "Reading Bridge",
    url: `assets/audio/cw25/cw25-day${dayIndex}.mp4`,
    audioUrl: `assets/audio/cw25/cw25-day${dayIndex}.mp4`,
    durationSeconds,
    summary,
    storyText,
    skillMatch,
    usedWords,
    keywords: uniqueWords([title, ...usedWords]),
    questions
  });
  return [
    story(
      0,
      "The Picture Day Plan",
      "Mia helps her class get ready for Picture Day by answering where everyone should stand.",
      "Setting, who/where/what, Picture Day, answer, where, now.",
      "Mia liked plans. On Picture Day, she held a small card with three questions: Who is in the photo? Where should we stand? What should we do now? The class walked into the hall, and everyone talked at once. Ben wanted the right side. A boy named Leo stood too far down the line. Ms. Ray smiled and asked, \"Who can help us solve this?\" Mia looked at the class photo sign. She gave a clear answer. \"Tall friends can stand in the back. Short friends can stand in front. Leo can stand here, where the tape makes a point.\" The class moved quietly. Now everyone could see the camera. Before the photo, Mia checked one more thing. \"Our favorite part is that we are together,\" she said. When the camera flashed, the class looked ready, kind, and proud.",
      ["answer", "where", "now", "right", "point", "favorite"],
      [
        cw25Question("Where does the story happen?", "in the school hall", ["at a zoo", "in a kitchen", "under a bridge"]),
        cw25Question("What problem does the class have?", "they do not know where to stand", ["they lose a book", "it starts snowing", "the camera breaks"]),
        cw25Question("How does Mia help?", "she gives a clear plan", ["she runs away", "she hides the sign", "she paints the wall"]),
        cw25Question("Which word describes Mia's answer?", "helpful", ["sleepy", "angry", "tiny"])
      ],
      46
    ),
    story(
      1,
      "The Boy Who Found a Better Way",
      "A boy uses a simple machine idea to move a heavy box safely.",
      "Problem and solution with simple machines, boy, found, how, solve.",
      "Owen found a heavy box near the classroom door. He pushed it, but it would not move. He pulled it, but it only made a loud sound on the floor. \"How can we move it safely?\" asked his teacher. Owen remembered the simple machines chart. A ramp can help things go up or down. A wheel can help things roll. A lever can help lift. Owen did not try to be strong. He tried to think like an engineer. He found a small cart with four wheels. Then he asked two friends to help him lift one side of the box. They slid the cart under it and walked slowly down the hall. The box moved without scratching the floor. Owen smiled because the answer was not to push harder. The answer was to solve the problem in a better way.",
      ["boy", "found", "how", "solve", "engineer", "down"],
      [
        cw25Question("What does Owen find?", "a heavy box", ["a green kite", "a lost dog", "a tiny hat"]),
        cw25Question("What does Owen use to move the box?", "a cart with wheels", ["a spoon", "a blanket", "a camera"]),
        cw25Question("Why does Owen think like an engineer?", "he wants to solve the problem safely", ["he wants to make noise", "he wants to hide", "he wants to eat lunch"]),
        cw25Question("What is the lesson?", "thinking can make hard work easier", ["never ask friends", "boxes should stay still", "only running helps"])
      ],
      42
    ),
    story(
      2,
      "The Ramp on the Playground",
      "Children design a ramp so a toy truck can roll down without tipping over.",
      "Sequence and engineering: engineer, design, down, out.",
      "At recess, Ana and Jay built a town with blocks. Their toy truck needed to go from the high bridge down to the road. First, Jay made a steep ramp. The truck rushed down and tipped over. Next, Ana made the ramp longer. The truck rolled, but it bumped into the wall. \"We need a better design,\" Ana said. They looked at the bridge, the road, and the open space out by the fence. Jay moved the road to the right. Ana added two blocks under the ramp so it would not shake. Then the truck rolled down, turned at the point, and came out safely by the pretend store. Their friends clapped. Ana did not call the first ramp bad. She said, \"It taught us what to change.\" Jay nodded. Engineers test, reflect, and try again.",
      ["engineer", "design", "down", "out", "right", "point"],
      [
        cw25Question("What are Ana and Jay building?", "a block town", ["a cake", "a moon", "a classroom door"]),
        cw25Question("What happens to the first ramp?", "the truck tips over", ["the truck flies away", "the ramp sings", "the blocks melt"]),
        cw25Question("What do they do after testing?", "they change the design", ["they stop playing", "they hide the truck", "they throw away the blocks"]),
        cw25Question("What do engineers do in this story?", "test, reflect, and try again", ["sleep, shout, and run", "guess and never look", "paint only blue"])
      ],
      44
    ),
    story(
      3,
      "A Voice for the Class Photo",
      "Nora shares an opinion about the class photo and gives a reason.",
      "Opinion and reason with voice, favorite, because.",
      "After Picture Day, the class looked at two photos. In the first photo, everyone smiled. In the second photo, the class held their favorite books. \"Which one should go on our board?\" asked Ms. Ray. Some children pointed to the first photo. Some pointed to the second. Nora wanted to use her voice, but she felt shy. She watched her friends give answers. Then she raised her hand. \"I choose the second photo because it shows what we love to read,\" she said. \"It tells a little story about our class.\" Ben nodded. \"I like that reason,\" he said. The class voted. The second photo won. Nora felt proud, not because everyone agreed at first, but because she shared an opinion with a reason. Her voice helped the class make a thoughtful choice.",
      ["voice", "favorite", "answer", "point", "watch", "because"],
      [
        cw25Question("What is the class choosing?", "which photo to put on the board", ["which soup to cook", "which dog to adopt", "which shoe to wear"]),
        cw25Question("Why does Nora choose the second photo?", "it shows what the class loves to read", ["it is bigger", "it has no children", "it is upside down"]),
        cw25Question("How does Nora use her voice?", "she shares an opinion with a reason", ["she sings a loud song", "she whispers nothing", "she reads the lunch menu"]),
        cw25Question("What does Nora feel at the end?", "proud", ["lost", "sleepy", "scared of books"])
      ],
      45
    ),
    story(
      4,
      "Watch the Wheel Turn",
      "A child observes a wheel carefully and writes details about how it works.",
      "Observation, details, watch, right, point, write.",
      "Luis had to write three details about a simple machine. He chose a wheel on the art cart. At first, he wrote, \"The wheel is round.\" Then he stopped. One detail was not enough. He sat down and watched the wheel turn. When the cart moved right, the wheel rolled smoothly. When the cart hit a bump, the wheel shook. Luis pointed to the middle part and asked, \"What is this called?\" His teacher said it was the axle. Luis wrote a new sentence: The wheel turns around an axle so the cart can move. Then he added: Wheels help heavy things move with less pushing. At share time, Luis read his details. His classmates understood how the wheel worked because Luis did more than look. He watched, asked, pointed, and wrote.",
      ["watch", "right", "point", "write", "how", "wheel"],
      [
        cw25Question("What simple machine does Luis study?", "a wheel", ["a pulley", "a screw", "a wedge"]),
        cw25Question("What does Luis do before writing better details?", "he watches the wheel turn", ["he closes his eyes", "he eats a snack", "he leaves the room"]),
        cw25Question("What is the middle part called?", "the axle", ["the photo", "the river", "the ribbon"]),
        cw25Question("Why do classmates understand Luis?", "he uses clear details", ["he writes nothing", "he draws a cow", "he runs fast"])
      ],
      44
    ),
    story(
      5,
      "The Cow and the Town Parade",
      "A playful parade story highlights ow/ou words in meaningful context.",
      "Diphthongs ow/ou with cow, town, now, out.",
      "The town parade was about to begin, but one cow stood in the middle of the road. \"How did she get out?\" asked Mayor Lee. The band waited. The children watched from the sidewalk. The cow looked down at a red flower and chewed slowly. \"Now what?\" said Jay. Mia had an idea. She found a bucket of water and set it beside the fence. The cow lifted her head. She walked toward the bucket, out of the road and through the open gate. The crowd clapped. The band began to play. Jay laughed and said, \"The cow was the first float!\" At the end of the parade, Mayor Lee gave Mia a paper ribbon. It said: Good problem solver. Mia smiled because she did not shout. She watched, thought, and helped.",
      ["cow", "town", "now", "out", "found", "how", "down"],
      [
        cw25Question("What blocks the parade?", "a cow", ["a train", "a snowman", "a giant book"]),
        cw25Question("Where is the cow standing?", "in the middle of the road", ["on the roof", "under a desk", "inside a backpack"]),
        cw25Question("How does Mia move the cow?", "she uses a bucket of water", ["she paints a sign", "she sings a song", "she builds a rocket"]),
        cw25Question("Why does Mia get a ribbon?", "she solves the problem calmly", ["she wins a race", "she eats the flower", "she drops the camera"])
      ],
      41
    ),
    story(
      6,
      "Celebrate the Small Fix",
      "A class reflects on a small repair and celebrates useful teamwork.",
      "Reflect, celebrate, weekly wrap-up, engineering problem solving.",
      "On Friday, the reading corner sign kept falling down. Tape did not hold it. A chair bumped it. The sign slid to the floor again and again. \"We can fix this,\" said Mei. The class looked at the wall, the string, and the tiny hook. Owen said the hook was too low. Nora said the string was too loose. Their teacher asked them to reflect before trying again. First, they named the problem. Next, they chose a better point on the wall. Then they tied the string tighter and lifted the sign. This time it stayed. The fix was small, but the class cheered. They did not celebrate a perfect sign. They celebrated careful thinking, helpful voices, and a design that worked. At the end of the day, the sign was still up, quietly telling everyone where good reading begins.",
      ["reflect", "celebrate", "solve", "design", "voice", "where"],
      [
        cw25Question("What keeps falling down?", "the reading corner sign", ["the clock", "a lunch tray", "a toy truck"]),
        cw25Question("What does the class do before trying again?", "they reflect on the problem", ["they run outside", "they close the books", "they turn off the lights"]),
        cw25Question("Why does the fix work?", "they choose a better point and tie the string tighter", ["they use more noise", "they paint the floor", "they hide the sign"]),
        cw25Question("What does the class celebrate?", "careful thinking and teamwork", ["a race", "a birthday cake", "a lost pencil"])
      ],
      47
    )
  ].filter(hasPlayableAudio);
}

function cw24ListeningStories() {
  return [
    normalizeListeningStory({
      dayIndex: 0,
      title: "The Best Dressed Crocodile",
      sourceName: "Storynory",
      url: "https://www.storynory.com/a-crocodile-goes-shopping/",
      audioUrl: "https://traffic.libsyn.com/blogrelations/A-Crocodile-Goes-Shopping-16.mp3",
      summary: "A crocodile shops for clothes, loses a handbag, and learns she can still be happy.",
      skillMatch: "Animal fantasy with problem/solution and usually/neat story language.",
      usedWords: ["problem", "neat"],
      keywords: ["crocodile", "shopping", "handbag", "happy"],
      questions: [
        networkStoryQuestion("What animal goes shopping?", "a crocodile", ["a penguin", "a rabbit", "a monkey"]),
        networkStoryQuestion("What problem happens in the cafe?", "someone takes her handbag", ["the moon falls", "a ship sinks", "a screw breaks"]),
        networkStoryQuestion("What does the crocodile learn?", "she can be happy without the handbag", ["she must never eat cake", "she should live on the moon", "she cannot shop"]),
        networkStoryQuestion("Which word best describes the story?", "funny", ["scary", "silent", "wordless"])
      ]
    }),
    normalizeListeningStory({
      dayIndex: 1,
      title: "Sleepy Star",
      sourceName: "Storynory",
      url: "https://www.storynory.com/sleepy-star/",
      audioUrl: "https://traffic.libsyn.com/blogrelations/sleepy-star-leaves-nest-16.mp3",
      summary: "A little star falls from the sky and becomes a child's bedtime treasure.",
      skillMatch: "Moon/sky imagination story for listening comprehension.",
      usedWords: ["dream", "moon"],
      keywords: ["star", "sky", "bedtime", "magic"],
      questions: [
        networkStoryQuestion("Where does the little star fall?", "to Earth", ["to the zoo", "into soup", "onto a boot"]),
        networkStoryQuestion("Who finds the star?", "a dog and Raffy", ["a crocodile", "three friends", "Joaquin"]),
        networkStoryQuestion("What does Raffy think the star is?", "magic", ["a screw", "a handle", "a cab"]),
        networkStoryQuestion("When is this story set?", "bedtime", ["lunchtime", "recess", "Friday test time"])
      ]
    }),
    normalizeListeningStory({
      dayIndex: 2,
      title: "The Wind and the Moon",
      sourceName: "Storynory",
      url: "https://www.storynory.com/the-wind-and-the-moon/",
      audioUrl: "https://traffic.libsyn.com/blogrelations/wind-moon-st.mp3",
      summary: "The wind tries to blow out the moon, but the moon keeps shining.",
      skillMatch: "Long oo word moon and poetic listening practice.",
      usedWords: ["moon", "blew"],
      keywords: ["wind", "moon", "blew", "poem"],
      questions: [
        networkStoryQuestion("What does the wind try to blow out?", "the moon", ["a candle", "a boot", "a zoo"]),
        networkStoryQuestion("Does the moon stop shining?", "no", ["yes", "only at noon", "only in soup"]),
        networkStoryQuestion("Which word from the story connects to this week's spelling?", "blew", ["scissors", "handle", "cab"]),
        networkStoryQuestion("What kind of text is this?", "a poem", ["a shopping list", "a recipe", "a test report"])
      ]
    }),
    normalizeListeningStory({
      dayIndex: 3,
      title: "The Moon",
      sourceName: "Storynory",
      url: "https://www.storynory.com/the-moon-poem/",
      audioUrl: "https://traffic.libsyn.com/secure/blogrelations/the-moon-poem-16.mp3",
      summary: "A short poem describes the moon shining over animals, streets, fields, and sleeping children.",
      skillMatch: "Long oo word moon and short audio poem.",
      usedWords: ["moon", "noon"],
      keywords: ["moon", "night", "animals", "poem"],
      questions: [
        networkStoryQuestion("What shines in the poem?", "the moon", ["a boot", "a zoo", "a cab"]),
        networkStoryQuestion("Which animals are out at night?", "cats, mice, dogs, and bats", ["penguins only", "crocodiles only", "no animals"]),
        networkStoryQuestion("What do flowers and children do at night?", "close their eyes", ["go shopping", "build a tool", "eat soup"]),
        networkStoryQuestion("What kind of text is this?", "a poem", ["a map", "a recipe", "a test"])
      ]
    }),
    normalizeListeningStory({
      dayIndex: 4,
      title: "The Monkey Who Loved Chocolate",
      sourceName: "Storynory",
      url: "https://www.storynory.com/the-monkey-who-loved-chocolate/",
      audioUrl: "https://traffic.libsyn.com/blogrelations/monkey-who-loved-chocolate2.mp3",
      summary: "Theo the monkey tastes chocolate at the zoo and creates a big problem in town.",
      skillMatch: "Zoo animal story with problem/solution listening practice.",
      usedWords: ["zoo", "problem"],
      keywords: ["monkey", "zoo", "chocolate", "problem"],
      questions: [
        networkStoryQuestion("What animal is Theo?", "a monkey", ["a crocodile", "a bear", "a star"]),
        networkStoryQuestion("Where does Jim first see Theo?", "at the zoo", ["on the moon", "in a kitchen", "at school"]),
        networkStoryQuestion("What does Theo love?", "chocolate", ["soup", "boots", "scissors"]),
        networkStoryQuestion("What problem does Theo cause?", "he steals chocolate", ["he loses a moon", "he builds a handle", "he washes shampoo"])
      ]
    }),
    normalizeListeningStory({
      dayIndex: 5,
      title: "The Bear's Breakfast",
      sourceName: "Storynory",
      url: "https://www.storynory.com/the-bears-breakfast/",
      audioUrl: "https://traffic.libsyn.com/blogrelations/Bears-breakfast-mx-storynory-16.mp3",
      summary: "A bear smells breakfast, goes into a house, and learns the little girl is scary.",
      skillMatch: "Short animal story with blew/shampoo vocabulary links.",
      usedWords: ["blew", "shampoo", "golden"],
      keywords: ["bear", "breakfast", "porridge", "girl"],
      questions: [
        networkStoryQuestion("What does the bear smell?", "breakfast", ["paint", "a moon", "a screw"]),
        networkStoryQuestion("What food does the bear like best?", "porridge", ["chocolate cake", "soup", "a sandwich"]),
        networkStoryQuestion("What does the bear use in the bathroom?", "shampoo", ["a screw", "a cab", "a handle"]),
        networkStoryQuestion("Why does the bear run away?", "he sees a little girl", ["he loses a boot", "the moon is too bright", "he cannot find soup"])
      ]
    }),
    normalizeListeningStory({
      dayIndex: 6,
      title: "Christmas in the Jungle",
      sourceName: "Storynory",
      url: "https://www.storynory.com/christmas-in-the-jungle/",
      audioUrl: "https://traffic.libsyn.com/blogrelations/christmas-in-the-jungle-v3.mp3",
      summary: "Monkey brings a strange kind of Christmas magic into the jungle.",
      skillMatch: "Animal fantasy with moon, golden, and problem/solution practice.",
      usedWords: ["moon", "golden", "problem"],
      keywords: ["monkey", "jungle", "snow", "gift"],
      questions: [
        networkStoryQuestion("Who is the main animal in the story?", "Monkey", ["Crocodile", "Sleepy Star", "Raffy"]),
        networkStoryQuestion("Where does the story happen?", "in the jungle", ["at a zoo", "on the moon", "in a department store"]),
        networkStoryQuestion("What does Monkey find on the mountain?", "snow", ["soup", "a screw", "a handbag"]),
        networkStoryQuestion("What lesson fits the story?", "try to be kind", ["take everything", "hide all tools", "never help friends"])
      ]
    })
  ].filter(hasPlayableAudio);
}

function suggestedListeningStoriesForPlan(plan = activePlan()) {
  const text = [
    plan.readingTitle,
    plan.readingArticleTitle,
    plan.phonicsFocus,
    plan.grammarFocus,
    ...(plan.masteryWords ?? []),
    ...((plan.weeklyVocabulary ?? []).map((item) => item.word))
  ].join(" ").toLowerCase();
  if (/joaquin|zoo/.test(text)) return cw24ListeningStories();
  return [];
}

function listeningStoriesForPlan(plan = activePlan()) {
  const existing = Array.isArray(plan.listeningStories) ? plan.listeningStories.map(normalizeListeningStory).filter(hasPlayableAudio) : [];
  const suggested = suggestedListeningStoriesForPlan(plan);
  if (existing.length >= 7) return existing;
  return suggested.length ? suggested : existing;
}

function listeningStoryDayIndex(date = new Date()) {
  return Math.min(Math.max(date.getDay(), 0), 6);
}

function networkListeningStoryForDate(plan = activePlan(), date = todayKey()) {
  const dateObject = new Date(`${date}T00:00:00`);
  const useCw26Fallback = shouldUseCw26StoryRewardAudio(dateObject) && !cw26StoryRewardAudioIsCurrent(plan.listeningStories ?? []);
  const useCw25Fallback = !useCw26Fallback && shouldUseCw25StoryRewardAudio(plan, dateObject) && !cw25StoryRewardAudioIsCurrent(plan.listeningStories ?? []);
  const stories = useCw26Fallback
    ? cw26ListeningStories()
    : useCw25Fallback
      ? cw25ListeningStories()
      : listeningStoriesForPlan(plan);
  if (!stories.length) return null;
  const dayIndex = listeningStoryDayIndex(dateObject);
  const story = stories.find((item) => Number(item.dayIndex) === dayIndex) ?? stories[dayIndex % stories.length];
  return story ? normalizeListeningStory(story) : null;
}

function buildDailyListeningStory(plan = activePlan(), date = todayKey(), learnerId = state.activeLearnerId, progress = activeProgress()) {
  const networkStory = networkListeningStoryForDate(plan, date);
  if (networkStory) {
    return {
      ...networkStory,
      weeklyTitle: plan.readingArticleTitle || plan.readingTitle || "this week's story",
      sourceType: "network"
    };
  }
  const title = plan.readingArticleTitle || plan.readingTitle || "This Week's Story";
  const selectedWords = selectStoryWords(plan, date, learnerId, progress);
  const { first, second, third } = storyWordPhrase(selectedWords);
  const variants = [
    {
      storyTitle: "The Quiet Blue",
      character: "Nora",
      setting: "the corner table by the window",
      problem: "Nora's blue paper tore right down the middle.",
      solution: "She used the tear as a river and painted around it.",
      text: [
        `Nora sat at the corner table by the window with one sheet of blue paper and a ${first} tucked near her elbow.`,
        `Outside, the morning was pale and still. Inside, everyone was busy making pictures for a hallway show.`,
        `Nora wanted her picture to feel wide, like a sky after rain. She pressed too hard with her brush, and the paper tore with a soft sigh.`,
        `For a moment, Nora did not move. The tear looked long and crooked. It looked like a mistake that had walked across her work.`,
        `Then her friend Leo leaned close and whispered, "It looks like water."`,
        `Nora looked again. The tear became a ${second} river. The torn edge became a hill, and the hill made her picture feel ${third}. She painted tiny houses beside it and one small bird above it.`,
        `When the bell rang, Nora held up her picture. It was not the sky she had planned. It was quieter, and it was better than before.`,
        `Nora smiled because a mistake had opened a little door.`
      ].join(" ")
    },
    {
      storyTitle: "The Day the Wall Listened",
      character: "Sam",
      setting: "the art room after lunch",
      problem: "Sam could not choose what to paint on the empty wall.",
      solution: "He listened to classmates and made one picture from many ideas.",
      text: [
        `After lunch, the art room smelled like paper, water, and clean brushes.`,
        `Sam stood in front of the empty wall. It was so large that his idea felt small.`,
        `He thought of a ${first}. He thought of a ${second}. He thought of one ${third} line across the bottom. Then he thought of nothing at all.`,
        `One child said, "Paint the sound of recess." Another said, "Paint the quiet before snow."`,
        `Sam closed his eyes. He heard shoes in the hall, a chair bump, and someone laughing behind a book.`,
        `Slowly, he painted a path that curved like a song. He added a bright window, a soft shadow, and a tiny mark near the bottom.`,
        `The wall did not look empty anymore. It looked as if it had been listening all day.`,
        `Sam stepped back and felt brave enough to add one more color.`
      ].join(" ")
    },
    {
      storyTitle: "A Pocket for the Wind",
      character: "Lena",
      setting: "the playground fence",
      problem: "Lena's paper kite would not fly.",
      solution: "She changes the tail and waits for a gentler wind.",
      text: [
        `Lena made a paper kite during choice time. She drew a small ${first} in the middle and tied a ${second} ribbon to the end.`,
        `At recess, she ran beside the playground fence. The kite hopped once, twisted twice, and dropped into the grass.`,
        `"Maybe it is too shy," Lena said.`,
        `She tried again. The kite tugged, bent, and fell near her shoe.`,
        `Lena sat down and watched the trees. The wind was not rushing anymore. It was tapping leaf after leaf, soft and ${second}.`,
        `She made the ribbon tail longer. She smoothed one ${third} corner. Then she waited.`,
        `When the wind came back, the kite rose just above her hands. It did not fly high, but it stayed up long enough for Lena to laugh.`,
        `She carried it inside like a secret she had learned from the air.`
      ].join(" ")
    },
    {
      storyTitle: "The Smallest Color",
      character: "June",
      setting: "a rainy classroom",
      problem: "June thinks her tiny dot of color is not enough.",
      solution: "The class builds a picture around the tiny dot.",
      text: [
        `Rain tapped the classroom windows, one small sound after another.`,
        `June had only a dot of yellow paint left on her brush. Everyone else had wide lines, tall trees, and shining suns.`,
        `"Mine is too small," June said.`,
        `Her teacher bent down. "Small things can still begin something."`,
        `June touched the paper once. The yellow dot sat there like a tiny lamp.`,
        `Milo added a ${first} beside it. Ana painted a soft ${third} path. Ben made a round ${second} that seemed to roll toward the light.`,
        `Soon the little dot was the middle of a whole town waking up after rain.`,
        `June looked at the page and felt her quiet idea grow.`
      ].join(" ")
    },
    {
      storyTitle: "When the Box Hummed",
      character: "Owen",
      setting: "the classroom rug",
      problem: "A plain cardboard box seems boring.",
      solution: "Owen and his friends imagine it into something useful and surprising.",
      text: [
        `There was a plain box on the classroom rug. It had no paint, no ribbon, and no sign.`,
        `Owen tapped it. The box made a low hum because something inside shifted.`,
        `"It sounds like a tiny machine," he said.`,
        `Maya peeked in and found string, paper circles, and one wooden ${first}.`,
        `The children could have pushed the box away, but Owen had a ${third} idea. He made a little stage. Maya made a moon. Theo made a ${second} road across the floor.`,
        `By reading time, the box was a theater. The paper moon rose, the string stars moved, and everyone leaned closer.`,
        `The box had not changed very much. The children had changed the way they looked at it.`,
        `That made the room feel bigger.`
      ].join(" ")
    }
  ];
  const variant = variants[seededValue(`${learnerId}:${date}:daily-story`) % variants.length];
  const storyText = variant.text;
  const questions = [
    makeChoiceQuestion(
      `Who is the main character in "${variant.storyTitle}"?`,
      variant.character,
      ["Mia", "Leo", "Marisol"],
      `${date}:${variant.storyTitle}:q1`
    ),
    makeChoiceQuestion(
      "Where does the story happen?",
      variant.setting,
      ["under the ocean", "inside a train station", "at the grocery store"],
      `${date}:${variant.storyTitle}:q2`
    ),
    makeChoiceQuestion(
      "What is the problem?",
      variant.problem,
      ["The class cannot find lunch.", "The teacher loses a hat.", "A dog runs into school."],
      `${date}:${variant.storyTitle}:q3`
    ),
    makeChoiceQuestion(
      "How is the problem solved?",
      variant.solution,
      ["Everyone gives up.", "The picture is thrown away.", "The room becomes dark."],
      `${date}:${variant.storyTitle}:q4`
    ),
    makeChoiceQuestion(
      "What is a lesson from the story?",
      "A new idea can grow from a problem.",
      ["Mistakes are always bad.", "Only big ideas matter.", "Friends should work alone."],
      `${date}:${variant.storyTitle}:q5`
    )
  ];
  return {
    storyTitle: variant.storyTitle,
    weeklyTitle: title,
    storyText,
    cover: storyCover(variant.storyTitle),
    questions,
    usedWords: selectedWords,
    keywords: uniqueWords([variant.character, "problem", "idea", "picture", first, second, third])
  };
}

function buildWeeklyStoryPractice(plan = activePlan()) {
  return buildDailyListeningStory(plan);
}

function retellPromptForDate(promptSets, date = todayKey()) {
  if (!promptSets.length) return [];
  const dayIndex = new Date(`${date}T00:00:00`).getDay();
  return promptSets[Math.min(dayIndex, promptSets.length - 1)] ?? promptSets[0];
}

function freeSpeakingPromptBank() {
  return [
    [
      {
        prompt: "What was your favorite food this week? Why?",
        hints: ["Name the food.", "Tell why you liked it.", "Use because."],
        keywords: ["favorite", "food", "because", "week"],
        referenceAnswer: "My favorite food this week was pizza because it tasted good.",
        strongerAnswer: "My favorite food this week was pizza because it was warm and cheesy. I ate it with my family."
      },
      {
        prompt: "What was the most interesting thing last week? Tell two details.",
        hints: ["Name the thing.", "Tell two details.", "Use first or then."],
        keywords: ["interesting", "week", "detail", "because"],
        referenceAnswer: "The most interesting thing was art class. I made a picture.",
        strongerAnswer: "The most interesting thing last week was art class. First, I chose bright colors. Then, I made a picture I liked."
      }
    ],
    [
      {
        prompt: "What is one thing you want to do this weekend? Why?",
        hints: ["Name the activity.", "Tell why.", "Use I want to..."],
        keywords: ["weekend", "want", "because", "activity"],
        referenceAnswer: "I want to play outside this weekend because it is fun.",
        strongerAnswer: "This weekend, I want to play outside with my family because I like moving my body and getting fresh air."
      },
      {
        prompt: "Tell one plan for tomorrow. What will you do first?",
        hints: ["Say tomorrow I will...", "Tell first.", "Use a complete sentence."],
        keywords: ["tomorrow", "will", "first", "plan"],
        referenceAnswer: "Tomorrow I will read a book. First, I will choose the book.",
        strongerAnswer: "Tomorrow I will read a book with my family. First, I will choose a story, and then I will talk about my favorite part."
      }
    ],
    [
      {
        prompt: "What is something you felt proud of this week?",
        hints: ["Name what you did.", "Tell why you felt proud.", "Use because."],
        keywords: ["proud", "did", "because", "week"],
        referenceAnswer: "I felt proud because I finished my work.",
        strongerAnswer: "I felt proud this week because I kept trying even when the work was hard. I finished it carefully."
      },
      {
        prompt: "Tell about something you learned this week.",
        hints: ["I learned...", "Tell one detail.", "Use a school word if you can."],
        keywords: ["learned", "week", "detail", "school"],
        referenceAnswer: "I learned a new word this week.",
        strongerAnswer: "This week I learned a new word and used it in a sentence. One detail I remember is that the word helped me explain my idea."
      }
    ],
    [
      {
        prompt: "Tell about a book, show, or game you liked this week.",
        hints: ["Name it.", "Tell one part you liked.", "Use because."],
        keywords: ["book", "show", "game", "liked", "because"],
        referenceAnswer: "I liked a book because it was funny.",
        strongerAnswer: "I liked a book this week because the characters were funny. My favorite part was when the problem was solved."
      },
      {
        prompt: "Choose one character or person you liked. What made them interesting?",
        hints: ["Name the character or person.", "Tell one detail.", "Use interesting because..."],
        keywords: ["character", "person", "interesting", "because"],
        referenceAnswer: "I liked the character because the character was kind.",
        strongerAnswer: "I liked the character because the character made a helpful choice. That was interesting because it changed what happened next."
      }
    ],
    [
      {
        prompt: "Tell about something funny that happened this week.",
        hints: ["What happened?", "Who was there?", "Why was it funny?"],
        keywords: ["funny", "happened", "because", "week"],
        referenceAnswer: "Something funny happened when we played a game.",
        strongerAnswer: "Something funny happened this week when we played a game. Everyone laughed because the ending was a surprise."
      },
      {
        prompt: "Tell about something you did with your family.",
        hints: ["Name the activity.", "Tell who was there.", "Tell one detail."],
        keywords: ["family", "together", "detail", "activity"],
        referenceAnswer: "I ate dinner with my family.",
        strongerAnswer: "I ate dinner with my family and talked about my day. One detail I remember is that we laughed together."
      }
    ],
    [
      {
        prompt: "What is one goal you have for next week?",
        hints: ["My goal is...", "Tell why.", "Use I will..."],
        keywords: ["goal", "next", "week", "will"],
        referenceAnswer: "My goal is to read every day next week.",
        strongerAnswer: "My goal for next week is to read every day. I will try because reading helps me learn more words."
      },
      {
        prompt: "What will help you do better next week?",
        hints: ["Tell one helper or habit.", "Use because.", "Say I can..."],
        keywords: ["next", "week", "help", "because"],
        referenceAnswer: "Practice will help me do better next week.",
        strongerAnswer: "Practice will help me do better next week because I can learn from my mistakes and try again."
      }
    ]
  ];
}

function rotatedFreeSpeakingPromptSet(dayIndex, plan = activePlan()) {
  const bank = freeSpeakingPromptBank();
  if (!bank.length) return [];
  const weekSeed = plan.weekId || activeWeekIdForDate();
  const weekNumber = Number(weekSeed.match(/cw(\d{2})/i)?.[1]);
  const pairStart = (Number.isFinite(weekNumber) ? weekNumber : seededValue(`${weekSeed}:free-speaking`)) % bank.length;
  const offset = dayIndex === 6 ? 1 : 0;
  return bank[(pairStart + offset) % bank.length];
}

function joaquinRetellPromptSets() {
  const keywords = ["Joaquin", "zoo", "object", "problem", "imagination", "animals"];
  return [
    [
      {
        prompt: "Who is the story about, and what does Joaquin imagine?",
        hints: ["Tell who.", "Tell what he imagines.", "Use one story word."],
        keywords,
        referenceAnswer: "The story is about Joaquin. He imagines a zoo with animals.",
        strongerAnswer: "The story is about Joaquin. He looks at objects and uses his imagination. He imagines animals in a zoo."
      },
      {
        prompt: "Where does Joaquin's idea take him?",
        hints: ["Name the place.", "Tell one thing he sees or imagines."],
        keywords,
        referenceAnswer: "Joaquin's idea takes him to a zoo in his imagination.",
        strongerAnswer: "Joaquin uses his imagination to make a zoo. He can think about animals and objects in a new way."
      }
    ],
    [
      {
        prompt: "Tell what happens first and next in Joaquin's Zoo.",
        hints: ["First...", "Next...", "Use Joaquin's name."],
        keywords,
        referenceAnswer: "First, Joaquin looks at objects. Next, he imagines they can be animals.",
        strongerAnswer: "First, Joaquin looks carefully at objects. Next, his imagination helps him see animals. Then the objects feel like part of a zoo."
      },
      {
        prompt: "Tell one important detail from the story.",
        hints: ["One detail is...", "This detail matters because..."],
        keywords,
        referenceAnswer: "One important detail is that Joaquin uses his imagination.",
        strongerAnswer: "One important detail is that Joaquin can look at an object and imagine an animal. This matters because imagination helps him make the zoo."
      }
    ],
    [
      {
        prompt: "What is the problem or challenge in Joaquin's Zoo?",
        hints: ["The problem is...", "Joaquin solves it by..."],
        keywords,
        referenceAnswer: "The problem is that Joaquin needs to see ordinary objects in a new way.",
        strongerAnswer: "The problem is that the objects are not real animals. Joaquin solves it by using his imagination to turn the objects into a zoo."
      },
      {
        prompt: "How does Joaquin respond to the problem?",
        hints: ["He looks...", "He imagines...", "Because..."],
        keywords,
        referenceAnswer: "Joaquin responds by using his imagination.",
        strongerAnswer: "Joaquin looks carefully and imagines animals. He responds with a good attitude because he can make a zoo in his mind."
      }
    ],
    [
      {
        prompt: "What does Joaquin show about imagination?",
        hints: ["Imagination means...", "Joaquin uses it to..."],
        keywords,
        referenceAnswer: "Joaquin shows that imagination can help you see new ideas.",
        strongerAnswer: "Joaquin shows that imagination can change how something looks in your mind. He uses imagination to see animals and make a zoo."
      },
      {
        prompt: "Choose one word to describe Joaquin. Why?",
        hints: ["I think Joaquin is...", "because..."],
        keywords,
        referenceAnswer: "I think Joaquin is creative because he imagines a zoo.",
        strongerAnswer: "I think Joaquin is creative because he sees more than ordinary objects. He uses his imagination to make animals and a zoo."
      }
    ],
    [
      {
        prompt: "Tell the end of Joaquin's Zoo in your own words.",
        hints: ["At the end...", "Joaquin learns or shows..."],
        keywords,
        referenceAnswer: "At the end, Joaquin's imagination helps him make a zoo.",
        strongerAnswer: "At the end, Joaquin shows that imagination is powerful. He can look at objects and think about animals in a zoo."
      },
      {
        prompt: "What lesson or big idea does Joaquin's Zoo show?",
        hints: ["The story shows...", "because..."],
        keywords,
        referenceAnswer: "The story shows that imagination can help you see new ideas.",
        strongerAnswer: "The story shows that imagination can change how we look at ordinary objects. Joaquin uses imagination to make a zoo."
      }
    ],
    rotatedFreeSpeakingPromptSet(5),
    rotatedFreeSpeakingPromptSet(6)
  ];
}

function mainStoryRetellPromptSets(plan = activePlan()) {
  const title = plan.readingArticleTitle || plan.readingTitle || "this week's story";
  const keywords = uniqueWords([title, ...(plan.readingSkills ?? []), ...weeklyVocabulary().slice(0, 6).map((item) => item.word)]);
  return [
    [
      {
        prompt: `Who or what is ${title} about?`,
        hints: ["Tell who or what.", "Tell where.", "Use a complete sentence."],
        keywords,
        referenceAnswer: `${title} is about the main character or idea in the story.`,
        strongerAnswer: `${title} is about the main character or idea. I can tell who is in the story and where it happens.`
      },
      {
        prompt: `What is the main idea of ${title}?`,
        hints: ["The story is mostly about...", "Tell one detail."],
        keywords,
        referenceAnswer: `${title} is mostly about one important story idea.`,
        strongerAnswer: `${title} is mostly about one important idea. I can support it with a detail from the story.`
      }
    ],
    [
      {
        prompt: `Tell what happens at the beginning of ${title}.`,
        hints: ["At the beginning...", "Tell one detail."],
        keywords,
        referenceAnswer: `At the beginning of ${title}, an important event starts the story.`,
        strongerAnswer: `At the beginning of ${title}, the character or idea is introduced. One detail helps me understand what will happen next.`
      },
      {
        prompt: `What happens next in ${title}?`,
        hints: ["Next...", "Then...", "Use a story detail."],
        keywords,
        referenceAnswer: `Next, another important event happens in ${title}.`,
        strongerAnswer: `Next, the story gives another important event. This detail helps the story move forward.`
      }
    ],
    [
      {
        prompt: `Tell one important detail from ${title}.`,
        hints: ["One detail is...", "This matters because..."],
        keywords,
        referenceAnswer: `One important detail from ${title} helps explain the story.`,
        strongerAnswer: `One important detail from ${title} matters because it helps me understand the character, problem, or idea.`
      },
      {
        prompt: `Use first, next, or then to retell part of ${title}.`,
        hints: ["First...", "Next...", "Then..."],
        keywords,
        referenceAnswer: `First or next, I can tell one part of ${title}.`,
        strongerAnswer: `First, I can tell what happened. Next, I can add another detail from ${title}.`
      }
    ],
    [
      {
        prompt: `What problem or challenge happens in ${title}?`,
        hints: ["The problem is...", "The character or idea changes because..."],
        keywords,
        referenceAnswer: `The problem or challenge in ${title} is an important part of the story.`,
        strongerAnswer: `The problem or challenge in ${title} helps the character or idea change. I can explain it with a story detail.`
      },
      {
        prompt: `How does someone respond to the problem in ${title}?`,
        hints: ["The character responds by...", "Because..."],
        keywords,
        referenceAnswer: `Someone responds to the problem by trying a solution.`,
        strongerAnswer: `Someone responds to the problem by making a choice or trying a solution. This response changes what happens next.`
      }
    ],
    [
      {
        prompt: `Tell the ending of ${title} in your own words.`,
        hints: ["At the end...", "Tell what changes."],
        keywords,
        referenceAnswer: `At the end of ${title}, the story reaches an ending.`,
        strongerAnswer: `At the end of ${title}, something changes or becomes clear. I can retell the ending with a detail.`
      },
      {
        prompt: `What lesson or big idea can you tell from ${title}?`,
        hints: ["The story shows...", "I think this because..."],
        keywords,
        referenceAnswer: `${title} shows an important lesson or big idea.`,
        strongerAnswer: `${title} shows an important lesson or big idea. I can explain my thinking with a story detail.`
      }
    ],
    rotatedFreeSpeakingPromptSet(5, plan),
    rotatedFreeSpeakingPromptSet(6, plan)
  ];
}

function hasMainStoryRetellEvidence(plan = activePlan()) {
  return Boolean(String(plan.readingArticleEvidence || "").trim());
}

function retellEvidenceLines(plan = activePlan()) {
  return splitRetellEvidence(plan.readingArticleEvidence || "").slice(0, 8);
}

function storyDetailOrFallback(value, fallback = "a detail from the story") {
  return String(value || "").trim().replace(/\s+/g, " ") || fallback;
}

function evidenceBasedRetellPromptSets(plan = activePlan()) {
  const title = plan.readingArticleTitle || plan.readingTitle || "this week's story";
  const evidence = retellEvidenceLines(plan);
  const characters = uniqueWords(plan.readingArticleCharacters ?? []).join(", ") || "the people or ideas in the story";
  const setting = storyDetailOrFallback(plan.readingArticleSetting, evidence[0] || "where the story happens");
  const events = uniqueWords([...(plan.readingArticleKeyEvents ?? []), ...evidence]).slice(0, 6);
  const firstEvent = storyDetailOrFallback(events[0], evidence[0] || "the first important part");
  const nextEvent = storyDetailOrFallback(events[1], evidence[1] || "the next important part");
  const detail = storyDetailOrFallback(events[2], evidence[2] || events[0] || "one important detail");
  const problem = storyDetailOrFallback(plan.readingArticleProblem, events.find((item) => /problem|challenge|need|try|hard|could/i.test(item)) || "the challenge in the story");
  const ending = storyDetailOrFallback(plan.readingArticleEnding, events[events.length - 1] || "the ending of the story");
  const bigIdea = storyDetailOrFallback(plan.readingArticleBigIdea, evidence[0] || "the big idea of the story");
  const keywords = uniqueWords([
    title,
    ...characters.split(",").map((item) => item.trim()),
    ...weeklyVocabulary().slice(0, 6).map((item) => item.word)
  ]);
  return [
    [
      {
        prompt: `Who or what is ${title} about?`,
        hints: ["Tell who or what.", "Tell where.", "Use a story detail."],
        keywords,
        referenceAnswer: `${title} is about ${characters}.`,
        strongerAnswer: `${title} is about ${characters}. It happens around ${setting}.`
      },
      {
        prompt: `Where does ${title} happen, or what topic does it teach?`,
        hints: ["Tell where or what topic.", "Use one detail from the story."],
        keywords,
        referenceAnswer: `One setting or topic detail is: ${setting}.`,
        strongerAnswer: `One important setting or topic detail is: ${setting}. This helps me understand ${title}.`
      }
    ],
    [
      {
        prompt: `Tell what happens first in ${title}.`,
        hints: ["Start with First...", "Use a real story detail."],
        keywords,
        referenceAnswer: `First, ${firstEvent}`,
        strongerAnswer: `First, ${firstEvent} This beginning detail helps the reader know what the story is about.`
      },
      {
        prompt: `What happens next in ${title}?`,
        hints: ["Start with Next...", "Tell another story detail."],
        keywords,
        referenceAnswer: `Next, ${nextEvent}`,
        strongerAnswer: `Next, ${nextEvent} This detail helps the story or information move forward.`
      }
    ],
    [
      {
        prompt: `Tell one important detail from ${title}.`,
        hints: ["One detail is...", "This matters because..."],
        keywords,
        referenceAnswer: `One important detail is: ${detail}.`,
        strongerAnswer: `One important detail is: ${detail}. This matters because it helps explain the story or topic.`
      },
      {
        prompt: `Use one weekly word to talk about ${title}.`,
        hints: ["Choose a weekly word.", "Use because if you can."],
        keywords,
        referenceAnswer: `I can use a weekly word to tell about this detail: ${detail}.`,
        strongerAnswer: `I can use a weekly word in a complete sentence about ${title}. My sentence connects the word to this detail: ${detail}.`
      }
    ],
    [
      {
        prompt: `What problem, question, or challenge appears in ${title}?`,
        hints: ["The problem or question is...", "Use a real detail."],
        keywords,
        referenceAnswer: `The problem, question, or challenge is: ${problem}.`,
        strongerAnswer: `The problem, question, or challenge is: ${problem}. I can explain it with a detail from the text.`
      },
      {
        prompt: `How does someone respond, solve, or explain the challenge in ${title}?`,
        hints: ["They respond by...", "Tell what changes or is explained."],
        keywords,
        referenceAnswer: `A response or explanation connects to this detail: ${nextEvent}.`,
        strongerAnswer: `A response or explanation connects to this detail: ${nextEvent}. This helps the reader understand what changes or what is learned.`
      }
    ],
    [
      {
        prompt: `Tell the ending or final important idea of ${title}.`,
        hints: ["At the end...", "Tell one final detail."],
        keywords,
        referenceAnswer: `At the end or final part, ${ending}`,
        strongerAnswer: `At the end or final part, ${ending} This final detail helps me retell the text clearly.`
      },
      {
        prompt: `What big idea can you tell from ${title}?`,
        hints: ["The big idea is...", "I think this because..."],
        keywords,
        referenceAnswer: `The big idea is: ${bigIdea}.`,
        strongerAnswer: `The big idea is: ${bigIdea}. I can support this idea with details from ${title}.`
      }
    ],
    rotatedFreeSpeakingPromptSet(5, plan),
    rotatedFreeSpeakingPromptSet(6, plan)
  ];
}

function poemSelectionsForPlan(plan = activePlan()) {
  return normalizePoemSelections(plan.poemSelections ?? []);
}

function poemSelectionForDay(plan = activePlan(), dayIndex = new Date(`${todayKey()}T00:00:00`).getDay()) {
  const poems = poemSelectionsForPlan(plan);
  if (!poems.length || dayIndex < 1 || dayIndex > 5) return null;
  return poems[dayIndex - 1] || null;
}

function hasPoemRetellEvidence(poem) {
  return Boolean(poem?.title && String(poem.evidence || "").trim());
}

function poemPromptSet(poem, plan = activePlan()) {
  const title = poem?.title || "today's poem";
  const evidence = storyDetailOrFallback(poem?.evidence, "the poem gives one clear picture or idea");
  const focus = splitEditableList(poem?.focus || "poetry, rhyme, rhythm, repetition, big idea").slice(0, 6);
  const author = poem?.author ? ` by ${poem.author}` : "";
  const keywords = uniqueWords([
    title,
    ...focus,
    ...weeklyVocabulary().slice(0, 6).map((item) => item.word)
  ]);
  const pictureHint = /shoe/i.test(title)
    ? "Think about the shoes, dreams, and steps."
    : /flying/i.test(title)
      ? "Think about the sky, sea, and questions."
      : /me x 2/i.test(title)
        ? "Think about the repeated times-two idea."
        : /row/i.test(title)
          ? "Think about rowing, rhythm, and the dream feeling."
          : /star/i.test(title)
            ? "Think about the star, night, and wish."
            : "Think about the picture the poem makes.";
  return [
    {
      prompt: `What is ${title}${author} mostly showing or saying?`,
      hints: ["The poem shows...", pictureHint, "Use one poem detail."],
      keywords,
      referenceAnswer: `${title} shows this idea: ${evidence}`,
      strongerAnswer: `${title} shows this idea: ${evidence} I can explain the poem in my own words and add one detail I noticed.`
    },
    {
      prompt: `What do you notice about the words, rhythm, feeling, or big idea in ${title}?`,
      hints: ["I notice...", "The repeated words or sounds help...", "My favorite part is..."],
      keywords,
      referenceAnswer: `I notice ${focus.join(", ") || "a poetry feature"} in ${title}.`,
      strongerAnswer: `I notice ${focus.join(", ") || "a poetry feature"} in ${title}. This helps the poem feel meaningful because it connects to this detail: ${evidence}`
    }
  ];
}

function normalizeRetellPrompt(prompt, retell) {
  const fallback = {
    prompt: "Tell one thing you remember from the story.",
    hints: ["First...", "Next...", "At the end..."],
    keywords: retell.keywords ?? [],
    referenceAnswer: "Tell one clear story detail in a complete sentence.",
    strongerAnswer: "Use two story details and a sequence word like first, next, or at the end."
  };
  if (prompt && typeof prompt === "object") {
    return {
      ...fallback,
      ...prompt,
      hints: Array.isArray(prompt.hints) ? prompt.hints : fallback.hints,
      keywords: Array.isArray(prompt.keywords) ? prompt.keywords : fallback.keywords
    };
  }
  return {
    ...fallback,
    prompt: String(prompt || "Tell one thing you remember from the story."),
  };
}

function buildWeeklyRetellPractice(plan = activePlan()) {
  if (isSummerRazDate()) return buildSummerRazRetellPractice();
  const title = plan.readingArticleTitle || plan.readingTitle || "this week's story";
  const dayIndex = new Date(`${todayKey()}T00:00:00`).getDay();
  const isPoetryPlan = normalizeWord(plan.readingArticleGenre) === "poetry" || poemSelectionsForPlan(plan).length > 0;
  if (isPoetryPlan) {
    const poem = poemSelectionForDay(plan, dayIndex);
    if (!poem) {
      const prompts = retellPromptForDate([rotatedFreeSpeakingPromptSet(0, plan), [], [], [], [], [], rotatedFreeSpeakingPromptSet(6, plan)]);
      return {
        weeklyTitle: "Poetry talk",
        prompts,
        sourceKey: `${plan.weekId || weekKey()}:${todayKey()}:poetry-free:${prompts.map((item) => item.prompt).join("|")}`,
        keywords: uniqueWords([...(plan.readingSkills ?? []), ...weeklyVocabulary().slice(0, 6).map((item) => item.word)]),
        retellSignals: ["because", "favorite", "noticed", "poem", "words", "feeling"]
      };
    }
    if (!hasPoemRetellEvidence(poem)) {
      return {
        weeklyTitle: poem.title,
        prompts: [],
        sourceKey: `${plan.weekId || weekKey()}:${todayKey()}:${poem.title}:missing-poem-evidence`,
        keywords: [],
        retellSignals: [],
        missingStoryEvidence: true,
        missingPoemEvidence: true
      };
    }
    const prompts = poemPromptSet(poem, plan);
    return {
      weeklyTitle: poem.title,
      prompts,
      sourceKey: `${plan.weekId || weekKey()}:${todayKey()}:${poem.title}:${prompts.map((item) => item.prompt).join("|")}`,
      keywords: uniqueWords([poem.title, ...(plan.readingSkills ?? []), ...weeklyVocabulary().slice(0, 6).map((item) => item.word)]),
      retellSignals: ["poem", "words", "repeat", "rhyme", "rhythm", "feeling", "big idea", "because"]
    };
  }
  if (dayIndex >= 5) {
    const prompts = retellPromptForDate([[], [], [], [], [], rotatedFreeSpeakingPromptSet(5, plan), rotatedFreeSpeakingPromptSet(6, plan)]);
    return {
      weeklyTitle: title,
      prompts,
      sourceKey: `${plan.weekId || weekKey()}:${todayKey()}:${title}:${prompts.map((item) => item.prompt).join("|")}`,
      keywords: uniqueWords([title, ...(plan.readingSkills ?? []), ...weeklyVocabulary().slice(0, 6).map((item) => item.word)]),
      retellSignals: ["because", "first", "next", "detail", "favorite", "learned"]
    };
  }
  if (!hasMainStoryRetellEvidence(plan)) {
    return {
      weeklyTitle: title,
      prompts: [],
      sourceKey: `${plan.weekId || weekKey()}:${todayKey()}:${title}:missing-main-story-pages`,
      keywords: [],
      retellSignals: [],
      missingStoryEvidence: true
    };
  }
  const promptSets = evidenceBasedRetellPromptSets(plan);
  const prompts = retellPromptForDate(promptSets);
  return {
    weeklyTitle: title,
    prompts,
    sourceKey: `${plan.weekId || weekKey()}:${todayKey()}:${title}:${prompts.map((item) => item.prompt).join("|")}`,
    keywords: uniqueWords([title, ...(plan.readingSkills ?? []), ...weeklyVocabulary().slice(0, 6).map((item) => item.word)]),
    retellSignals: ["first", "next", "then", "because", "beginning", "middle", "end", "problem", "solution", "learns", "changes"]
  };
}

function buildSummerRazRetellPractice() {
  const dayProgress = ensureAdventureDayProgress(activeProgress());
  const articles = razAssignmentsFromProgress(dayProgress);
  let selectedArticle = articles.find((article) => article.id === dayProgress.razRetell.selectedArticleId) || null;
  if (!selectedArticle && articles.length === 1) {
    selectedArticle = articles[0];
    dayProgress.razRetell.selectedArticleId = selectedArticle.id;
    dayProgress.razRetell.selectedAt = new Date().toISOString();
  }
  if (selectedArticle) queueRazArticleAnalysis([selectedArticle]);
  const details = selectedArticle ? store.razArticleDetails?.[selectedArticle.id] ?? {} : {};
  const informational = details.genre === "informational";
  const evidence = splitRetellEvidence(details.readingArticleEvidence || "");
  const characters = (details.readingArticleCharacters ?? []).join(", ");
  const keyEvents = details.readingArticleKeyEvents?.length ? details.readingArticleKeyEvents : evidence.slice(0, 3);
  const fingers = informational
    ? [
        { label: "1. Topic", prompt: `What is ${selectedArticle?.title || "the book"} mostly about?`, support: details.readingArticleBigIdea || evidence[0] || "Name the main topic." },
        { label: "2. Place or Time", prompt: "Where or when does this information happen?", support: details.readingArticleSetting || "Use a place, time, or text feature from the book." },
        { label: "3. Main Question", prompt: "What important question does the book answer?", support: details.readingArticleProblem || "Say what you wanted to find out." },
        { label: "4. Key Details", prompt: "Tell two or three important facts in order.", support: keyEvents.join(" ") || "Use first, next, and then." },
        { label: "5. What I Learned", prompt: "What did you learn at the end?", support: details.readingArticleEnding || details.readingArticleBigIdea || "Finish with the most important new idea." }
      ]
    : [
        { label: "1. Characters", prompt: "Who is in the story?", support: characters || "Name the main character or characters." },
        { label: "2. Setting", prompt: "Where and when does the story happen?", support: details.readingArticleSetting || "Name the place and time if you know them." },
        { label: "3. Problem", prompt: "What problem or challenge happens?", support: details.readingArticleProblem || "Tell what the character wants or needs." },
        { label: "4. Key Events", prompt: "What happens first, next, and then?", support: keyEvents.join(" ") || "Tell the most important events in order." },
        { label: "5. Solution or Ending", prompt: "How is the problem solved, or how does the story end?", support: details.readingArticleEnding || "Tell the ending and what the character learns." }
      ];
  const prompt = {
    prompt: `Retell ${selectedArticle?.title || "your chosen RAZ book"} in one recording. Use all five fingers in order.`,
    hints: fingers.map((finger) => finger.label),
    keywords: uniqueWords([...(details.readingArticleCharacters ?? []), ...keyEvents.flatMap((item) => String(item).match(/[A-Za-z']{4,}/g) ?? [])]).slice(0, 8),
    referenceAnswer: fingers.map((finger) => finger.support).join(" "),
    strongerAnswer: "Use complete sentences and sequence words such as first, next, then, and finally."
  };
  return {
    weeklyTitle: selectedArticle?.title || "Summer RAZ retell",
    sourceKey: `raz-five-finger:${todayKey()}:${selectedArticle?.id || "choose"}`,
    sourceType: "raz-five-finger",
    articles,
    selectedArticle,
    fingers,
    prompts: selectedArticle ? [prompt] : [],
    keywords: prompt.keywords,
    retellSignals: informational ? ["topic", "first", "next", "learned", "because"] : ["first", "next", "then", "problem", "solution", "ending"],
    needsArticleSelection: articles.length > 1 && !selectedArticle,
    missingStoryEvidence: articles.length === 0
  };
}

function recordStoryWordExposure(words, progress = activeProgress(), plan = activePlan()) {
  const exposureKey = storyWordExposureKey(plan);
  if (!progress.storyWordExposure) progress.storyWordExposure = {};
  const current = { ...(progress.storyWordExposure[exposureKey] ?? {}) };
  words.forEach((word) => {
    const key = normalizeWord(word);
    if (!key) return;
    current[key] = (current[key] ?? 0) + 1;
  });
  progress.storyWordExposure[exposureKey] = current;
}

function ensureReadingQuizSession() {
  const progress = activeProgress();
  const date = todayKey();
  const story = buildDailyListeningStory(activePlan(), date, state.activeLearnerId, progress);
  const sourceKey = `${activePlan().weekId}:${date}:${story.storyTitle}:${story.questions.map((item) => item.prompt).join("|")}`;
  if (!progress.readingQuizSessions) progress.readingQuizSessions = {};
  const saved = progress.readingQuizSessions[date];
  if (saved?.sourceKey === sourceKey) return saved;
  recordStoryWordExposure(story.usedWords ?? [], progress);
  progress.readingQuizSessions[date] = {
    date,
    sourceKey,
    story,
    currentIndex: 0,
    answers: [],
    completed: false,
    score: 0,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  saveStore();
  return progress.readingQuizSessions[date];
}

function currentListeningStorySession() {
  return ensureReadingQuizSession();
}

function currentListeningStory() {
  return currentListeningStorySession().story ?? buildDailyListeningStory();
}

function ensureStoryRewardQuizSession() {
  const progress = activeProgress();
  const date = todayKey();
  const story = networkListeningStoryForDate(activePlan(), date);
  if (!story || !hasPlayableAudio(story)) return null;
  const sourceKey = `${activePlan().weekId}:${date}:story-reward:${story.storyTitle}:${story.audioUrl}:${story.questions.map((item) => item.prompt).join("|")}`;
  if (!progress.storyRewardQuizSessions) progress.storyRewardQuizSessions = {};
  const saved = progress.storyRewardQuizSessions[date];
  if (saved?.sourceKey === sourceKey) return saved;
  progress.storyRewardQuizSessions[date] = {
    date,
    sourceKey,
    story,
    currentIndex: 0,
    answers: [],
    completed: false,
    score: 0,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  saveStore();
  return progress.storyRewardQuizSessions[date];
}

function currentQuizSessionForActivePractice() {
  return state.activePracticeId === "story-reward"
    ? ensureStoryRewardQuizSession()
    : currentListeningStorySession();
}

function grammarChoice(prompt, answer, distractors, seed, skill) {
  return { ...makeChoiceQuestion(prompt, answer, distractors, seed), skill };
}

function grammarReviewSkills(progress = activeProgress()) {
  return Object.entries(progress.grammarSkillProfile ?? {})
    .filter(([, stat]) => (stat.missed ?? 0) > (stat.correct ?? 0))
    .sort((a, b) => (b[1].missed ?? 0) - (a[1].missed ?? 0))
    .map(([skill]) => skill);
}

function grammarQuestionForSkill(skill, seed, plan = activePlan()) {
  const title = plan.readingArticleTitle || plan.readingTitle || "the story";
  const topic = plan.uoiTopic || "class";
  const contractions = plan.grammarPracticeWords?.filter((word) => word.includes("'")) ?? [];
  const contraction = contractions[seededValue(`${seed}:contraction`) % Math.max(contractions.length, 1)] || "can't";
  const contractionExpanded = {
    "isn't": "is not",
    "that's": "that is",
    "we'll": "we will",
    "you'll": "you will",
    "can't": "can not",
    "didn't": "did not",
    "i'll": "I will",
    "i'm": "I am",
    "it's": "it is",
    "wasn't": "was not"
  }[normalizeWord(contraction)] ?? "two words joined together";
  const contractionSentence = {
    "isn't": "This isn't a real zoo.",
    "that's": "That's a clever idea.",
    "we'll": "We'll read the story after lunch.",
    "you'll": "You'll hear the word in a sentence.",
    "can't": "Joaquin can't stop imagining animals.",
    "didn't": "The class didn't give up.",
    "i'll": "I'll try the next question.",
    "i'm": "I'm ready to read.",
    "it's": "It's time for grammar practice.",
    "wasn't": "The answer wasn't hard after we read it."
  }[normalizeWord(contraction)] ?? `${displayWord(contraction)} helps make two words shorter.`;
  const builders = {
    comparing: () => grammarChoice("Choose the sentence that compares two things.", "This ramp is safer than that ramp.", ["This ramp is safe than that ramp.", "This ramp safest than that ramp.", "This ramp safer that ramp."], `${seed}:compare`, "comparing"),
    informational: () => grammarChoice(`Which sentence sounds like an informational text about ${topic}?`, `A ramp is a tool that helps people move things.`, ["The ramp danced under the moon.", "Once a ramp became a dragon.", "Wow, ramps are funny!"], `${seed}:info`, "informational"),
    biography: () => grammarChoice("Which sentence could belong in a biographical essay?", "Joaquin used his imagination to solve a problem.", ["The zoo flew into space.", "A magic hat ate lunch.", "Blue is my favorite snack."], `${seed}:bio`, "biography"),
    prefix: () => grammarChoice("Choose the word that means to do again.", "redo", ["untie", "cover", "happy"], `${seed}:prefix`, "prefix"),
    contraction: () => grammarChoice(`In this sentence, what does ${contraction} mean? "${contractionSentence}"`, contractionExpanded, ["was not", "they are", "very quickly"], `${seed}:contraction`, "contraction"),
    complete: () => grammarChoice("Choose the complete sentence.", `The class read ${title}.`, [`Read ${title}.`, "The class", `${title} with a book`], `${seed}:complete`, "complete-sentence"),
    punctuation: () => grammarChoice("Choose the sentence with correct capitalization and end mark.", "Joaquin imagines a zoo.", ["joaquin imagines a zoo.", "Joaquin imagines a zoo", "joaquin imagines a zoo"], `${seed}:punctuation`, "punctuation"),
    expand: () => grammarChoice("Choose the stronger sentence.", "Joaquin looked carefully because he wanted a new idea.", ["Joaquin looked.", "Carefully because idea.", "Wanted a new idea Joaquin."], `${seed}:expand`, "sentence-expand"),
    adjective: () => grammarChoice("Which word describes the noun in this sentence? The quiet classroom was ready.", "quiet", ["classroom", "was", "ready"], `${seed}:adjective`, "adjective"),
    adverb: () => grammarChoice("Which word tells how the child read? The child read softly.", "softly", ["child", "read", "the"], `${seed}:adverb`, "adverb")
  };
  return (builders[skill] ?? builders.complete)();
}

function grammarQuestions(plan = activePlan(), date = todayKey(), progress = activeProgress()) {
  const focus = `${plan.grammarFocus || ""} ${plan.readingArticleSummary || ""}`.toLowerCase();
  const skills = [];
  if (focus.includes("compar")) {
    skills.push("comparing");
  }
  if (focus.includes("informational")) {
    skills.push("informational", "biography");
  }
  if (focus.includes("prefix") || plan.grammarPracticeWords?.some((word) => ["re-", "un-", "redo", "untie", "uncover", "unhappy"].includes(normalizeWord(word)))) {
    skills.push("prefix");
  }
  const contractions = plan.grammarPracticeWords?.filter((word) => word.includes("'")) ?? [];
  if (contractions.length) skills.push("contraction");
  if (focus.includes("adverb")) skills.push("adverb");
  if (focus.includes("adjective")) skills.push("adjective");
  skills.push("complete", "punctuation", "expand");
  const reviewSkills = grammarReviewSkills(progress).slice(0, 2);
  const orderedSkills = uniqueWords([
    ...reviewSkills,
    ...seededShuffle(uniqueWords(skills), `${state.activeLearnerId}:${date}:grammar-skills`)
  ]).slice(0, 5);
  return orderedSkills.map((skill, index) => grammarQuestionForSkill(skill, `${state.activeLearnerId}:${date}:${index}`, plan));
}

function ensureGrammarSession() {
  const progress = activeProgress();
  const date = todayKey();
  if (!progress.grammarSkillProfile) progress.grammarSkillProfile = {};
  const sourceKey = `grammar-feedback-v3:${activePlan().weekId}:${date}:${activePlan().grammarFocus}:${(activePlan().grammarPracticeWords ?? []).map(normalizeWord).join("|")}`;
  if (!progress.grammarSessions) progress.grammarSessions = {};
  const saved = progress.grammarSessions[date];
  if (saved?.sourceKey === sourceKey) return saved;
  const questions = grammarQuestions(activePlan(), date, progress);
  progress.grammarSessions[date] = {
    date,
    sourceKey,
    questions,
    currentIndex: 0,
    answers: [],
    completed: false,
    score: 0,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  saveStore();
  return progress.grammarSessions[date];
}

function ensureSpellingSession() {
  const progress = activeProgress();
  const now = new Date();
  const date = todayKey(now);
  const key = practiceKey("listen", date);
  if (!progress.spellingSessions) progress.spellingSessions = {};
  if (!progress.weeklyExposure) progress.weeklyExposure = {};
  if (!progress.wrongWordProfile) progress.wrongWordProfile = {};
  const sundayReview = isSundaySpellingReview(now);
  const weeklyWords = sundayReview ? [] : weeklyNewWords();
  const sourceWords = spellingSourceWordsForDate(progress, now);
  const sourceKey = sundayReview
    ? `${SPELLING_SELECTION_VERSION}:sunday-review:${state.activeLearnerId}:${date}`
    : `${SPELLING_SELECTION_VERSION}:weekly:${activePlan().weekId}:${weeklyWords.map(normalizeWord).join("|")}`;
  const saved = progress.spellingSessions[date];
  if (saved?.words?.length && (saved.sourceKey === sourceKey || (sundayReview && saved.mode === "sunday-review"))) return saved;

  const weeklyKeys = new Set(weeklyWords.map(normalizeWord));
  const week = weekKey();
  const exposure = { ...(progress.weeklyExposure[week] ?? {}) };
  const target = dailySpellingTarget(progress, now);
  if (sundayReview) {
    const words = seededShuffle(sourceWords, `${state.activeLearnerId}:${date}:sunday-review`).slice(0, target);
    progress.spellingSessions[date] = {
      date,
      weekKey: week,
      mode: "sunday-review",
      words,
      currentIndex: 0,
      attemptsForCurrent: 0,
      completedWords: [],
      results: [],
      currentAnswer: "",
      completed: false,
      sourceKey,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    if (!progress.practiceStats) progress.practiceStats = {};
    progress.practiceStats[key] = {
      ...defaultPracticeStat(),
      ...(progress.practiceStats[key] ?? {}),
      completedItems: progress.practiceStats[key]?.completedItems ?? 0
    };
    saveStore();
    return progress.spellingSessions[date];
  }
  const minNewTarget = Math.min(
    weeklyWords.length,
    Math.max(
      Math.ceil(target * MIN_DAILY_SPELLING_NEW_WORD_RATIO),
      requiredNewWordsForThursdayCoverage(weeklyWords, exposure)
    )
  );
  const dueNewWords = [];
  const priorityWords = spellingPriorityWords(activePlan());
  for (let count = 0; count < WEEKLY_NEW_WORD_EXPOSURE_GOAL; count += 1) {
    const dueAtCount = weeklyWords.filter((word) => (exposure[normalizeWord(word)] ?? 0) === count);
    dueNewWords.push(...priorityWords.filter((word) => dueAtCount.map(normalizeWord).includes(normalizeWord(word))));
  }
  const newWordCandidates = uniqueWords([
    ...dueNewWords,
    ...priorityWords.filter((word) => !dueNewWords.map(normalizeWord).includes(normalizeWord(word)))
  ]);
  const selectedNewWords = newWordCandidates.slice(0, Math.min(target, minNewTarget));
  const reviewWords = seededShuffle(spellingReviewWords(progress), `${state.activeLearnerId}:${date}:review`);
  const reviewSlots = Math.max(0, target - selectedNewWords.length);
  const selectedKeys = new Set(selectedNewWords.map(normalizeWord));
  const restNewWords = weeklyWords.filter((word) => !selectedKeys.has(normalizeWord(word)));
  const candidates = uniqueWords([
    ...selectedNewWords,
    ...reviewWords.slice(0, reviewSlots),
    ...seededShuffle(restNewWords, `${state.activeLearnerId}:${date}:rest`)
  ]);
  const words = candidates.slice(0, target);

  words.forEach((word) => {
    const normalized = normalizeWord(word);
    if (weeklyKeys.has(normalized)) exposure[normalized] = (exposure[normalized] ?? 0) + 1;
  });
  progress.weeklyExposure[week] = exposure;
  progress.spellingSessions[date] = {
    date,
    weekKey: week,
    words,
    currentIndex: 0,
    attemptsForCurrent: 0,
    completedWords: [],
    results: [],
    currentAnswer: "",
    completed: false,
    sourceKey,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  if (!progress.practiceStats) progress.practiceStats = {};
  progress.practiceStats[key] = {
    ...defaultPracticeStat(),
    ...(progress.practiceStats[key] ?? {}),
    completedItems: progress.practiceStats[key]?.completedItems ?? 0
  };
  saveStore();
  return progress.spellingSessions[date];
}

function spellingSessionCompletedCount(session) {
  const words = session?.words ?? [];
  if (!words.length) return 0;
  const wordKeys = new Set(words.map(normalizeWord));
  const completedKeys = new Set(
    (session.completedWords ?? [])
      .map(normalizeWord)
      .filter((word) => wordKeys.has(word))
  );
  return Math.min(completedKeys.size, words.length);
}

function spellingSessionIsComplete(session) {
  const words = session?.words ?? [];
  return Boolean(words.length && (session?.completed || spellingSessionCompletedCount(session) >= words.length));
}

function centerProgress(center) {
  const stat = practiceStat(center.id);
  const spellingSession = center.id === "listen" ? ensureSpellingSession() : null;
  const goal = center.id === "listen"
    ? (spellingSession?.words?.length || dailySpellingTarget())
    : center.id === "sight"
      ? (ensureSightSession().activities?.length || ensureSightSession().words?.length || 0)
      : center.completionGoal || 1;
  const completedItems = Math.min(
    center.id === "listen"
      ? Math.max(stat.completedItems || 0, spellingSessionCompletedCount(spellingSession))
      : stat.completedItems || 0,
    goal
  );
  const complete = goal > 0 && (center.id === "listen" ? spellingSessionIsComplete(spellingSession) || completedItems >= goal : completedItems >= goal);
  const percent = goal ? Math.round((completedItems / goal) * 100) : 0;
  return {
    ...stat,
    goal,
    completedItems,
    complete,
    percent,
    status: complete ? "Complete" : completedItems > 0 || stat.attempts > 0 ? "In progress" : "Not started",
    label: complete ? `Complete · ${formatDuration(stat.secondsSpent || 0)}` : `${completedItems}/${goal} ${center.progressUnit || "items"}`
  };
}

function completedCenterIds() {
  return missionCenters
    .filter((center) => dailyMissionCenterIds.includes(center.id))
    .filter((center) => centerProgress(center).complete)
    .map((center) => center.id);
}

function recordPracticeAction(centerId, { correct = false, items = 1, seconds = 0, render = true } = {}) {
  const progress = activeProgress();
  if (!progress.practiceStats) progress.practiceStats = {};
  const key = practiceKey(centerId);
  const current = { ...defaultPracticeStat(), ...(progress.practiceStats[key] ?? {}) };
  const now = new Date().toISOString();
  const spellingSession = centerId === "listen" ? progress.spellingSessions?.[todayKey()] : null;
  const goal = centerId === "listen"
    ? (spellingSession?.words?.length || dailySpellingTarget(progress))
    : centerId === "sight"
      ? (ensureSightSession().activities?.length || ensureSightSession().words?.length || 0)
      : missionCenters.find((center) => center.id === centerId)?.completionGoal || 1;
  const completedItems = Math.max(current.completedItems, Math.min((current.completedItems || 0) + items, goal));
  progress.practiceStats[key] = {
    ...current,
    attempts: current.attempts + 1,
    correct: current.correct + (correct ? 1 : 0),
    completedItems,
    secondsSpent: current.secondsSpent + (timedPracticeIsActive(centerId) ? 0 : seconds),
    startedAt: current.startedAt || now,
    lastUpdated: now
  };
  if (completedItems >= goal && timedPracticeIsActive(centerId)) {
    stopPracticeTimer({ save: false });
  }
  if (completedItems >= goal && legacyCenterAdventureMap[centerId]) {
    completeAdventureLevel(legacyCenterAdventureMap[centerId], { render: false });
  }
  const completedIds = completedCenterIds();
  progress.centerCompletions[todayKey()] = completedIds;
  if (completedIds.length >= dailyMissionCenterIds.length && !progress.completedDays.includes(todayKey())) {
    progress.completedDays.push(todayKey());
  }
  saveStore();
  if (render) renderAll();
}

function clickableText(text) {
  return text.replace(/[A-Za-z][A-Za-z']*/g, (word) => {
    const key = normalizeWord(word);
    const weeklyKnown = activePlan().masteryWords?.map(normalizeWord).includes(key) || activePlan().sightWords?.map(normalizeWord).includes(key);
    if (!wordHelp[key] && !weeklyKnown) return word;
    return `<button class="word-link" type="button" data-word="${key}">${word}</button>`;
  });
}

function tokenizePlan() {
  const plan = activePlan();
  const raw = [
    plan.readingTitle,
    plan.phonicsFocus,
    plan.grammarFocus,
    plan.uoiTopic,
    ...(plan.masteryWords ?? []),
    ...(plan.sightWords ?? []),
    ...(plan.epicIdeas ?? [])
  ].join(" ");
  return raw
    .toLowerCase()
    .replace(/[^a-z0-9' ]/g, " ")
    .split(/\s+/)
    .filter((token) => token.length > 2);
}

function allBookEntries() {
  const parentBooks = (store.parentBooks ?? []).map((book) => ({
    ...book,
    source: book.source || "Parent shelf",
    sourceUrl: book.sourceUrl || "https://www.getepic.com/",
    topicTags: book.topicTags ?? [],
    skillTags: book.skillTags ?? [],
    searchTerms: book.searchTerms ?? [book.title]
  }));
  return [...parentBooks, ...epicSeedLibrary, ...epicSearchLibrary];
}

function scoreBook(book) {
  const tokens = tokenizePlan();
  const haystack = [
    book.title,
    book.author,
    book.format,
    ...(book.topicTags ?? []),
    ...(book.skillTags ?? []),
    ...(book.searchTerms ?? [])
  ].join(" ").toLowerCase();
  const directMatches = tokens.filter((token) => haystack.includes(token)).length;
  const parentBoost = book.source === "Parent shelf" ? 8 : 0;
  const phonicsBoost = /phonics|vowel|syllable|contraction|decode/.test(activePlan().phonicsFocus.toLowerCase()) &&
    (book.skillTags ?? []).some((tag) => /phonics|vowel|syllable|contraction|decode/.test(tag))
    ? 6
    : 0;
  const scienceBoost = /machine|push|pull|ramp|lever|wheel|architecture|building|design/.test(activePlan().uoiTopic.toLowerCase()) &&
    (book.topicTags ?? []).some((tag) => /machine|push|pull|ramp|lever|wheel|architecture|building|design/.test(tag))
    ? 6
    : 0;
  return directMatches + parentBoost + phonicsBoost + scienceBoost;
}

function recommendedBooks() {
  const scored = allBookEntries()
    .map((book) => ({ ...book, score: scoreBook(book) }))
    .sort((a, b) => b.score - a.score || a.title.localeCompare(b.title));
  const fallback = scored.slice(0, 3);
  const matched = scored.filter((book) => book.score > 0).slice(0, 3);
  const picks = matched.length ? matched : fallback;
  return ["Best Match", "Easy Win", "Stretch"].map((label, index) => ({
    label,
    ...(picks[index] ?? picks[0])
  }));
}

function primarySearchTerm(book) {
  return book.searchTerms?.[0] ?? book.title;
}

function parseTags(value) {
  return uniqueWords(String(value || "").split(/,|\n/).map((tag) => tag.toLowerCase()));
}

function materialExtension(fileName) {
  return fileName.split(".").pop()?.toLowerCase() ?? "";
}

function safeStorageName(fileName) {
  const cleaned = fileName
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
  return cleaned || `material-${Date.now()}`;
}

function createId() {
  return crypto.randomUUID?.() ?? `material-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function fileTypeLabel(material) {
  if (material.mimeType?.startsWith("image/")) return "image";
  if (material.mimeType === "application/pdf") return "pdf";
  if (["xlsx", "xls", "csv"].includes(material.fileType)) return "excel";
  return material.fileType || "file";
}

function materialReadyForWeeklyPlan(material = {}) {
  return material.materialKind === "weekly_plan" && weeklyPlanDataHasContent(material.weeklyPlanData);
}

function materialReadyForMainStory(material = {}) {
  return material.materialKind === "main_story_pages" && Boolean(material.storyDetails?.readingArticleEvidence);
}

function materialParseLabel(material = {}) {
  if (materialReadyForMainStory(material)) return "Ready for Level 2 retell";
  if (material.parseStatus === "parsed_weekly_plan" || materialReadyForWeeklyPlan(material)) return "Ready to apply to weekly plan";
  if (material.parseStatus === "needs_review_ocr") return "Needs OCR/manual review";
  if (material.parseStatus === "parsed") return "Parsed";
  if (material.parseStatus === "parsing") return "Parsing";
  if (material.parseStatus === "failed") return "Parsing failed";
  return material.parseStatus || "Saved";
}

function validateMaterialFile(file) {
  if (!file) return "Choose a file first.";
  const extension = materialExtension(file.name);
  if (!ALLOWED_MATERIAL_EXTENSIONS.includes(extension)) {
    return "Supported formats: JPG, PNG, WEBP, PDF, XLSX, XLS, CSV.";
  }
  if (file.size > MAX_MATERIAL_SIZE) {
    return "File is too large. Maximum size is 25MB.";
  }
  if (file.type && !ALLOWED_MATERIAL_MIME_TYPES.includes(file.type)) {
    return "This file type is not supported.";
  }
  return "";
}

function openMaterialDb() {
  if (materialDbPromise) return materialDbPromise;
  materialDbPromise = new Promise((resolve, reject) => {
    const request = indexedDB.open(MATERIAL_DB_NAME, MATERIAL_DB_VERSION);
    request.addEventListener("upgradeneeded", () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(MATERIAL_STORE)) {
        const storeObject = db.createObjectStore(MATERIAL_STORE, { keyPath: "id" });
        storeObject.createIndex("createdAt", "createdAt");
        storeObject.createIndex("materialKind", "materialKind");
      }
      if (!db.objectStoreNames.contains(SIGHT_WORD_STORE)) {
        const sightStore = db.createObjectStore(SIGHT_WORD_STORE, { keyPath: "word" });
        sightStore.createIndex("grade", "grade");
        sightStore.createIndex("source", "source");
      }
    });
    request.addEventListener("success", () => resolve(request.result));
    request.addEventListener("error", () => reject(request.error));
  });
  return materialDbPromise;
}

async function materialStore(mode = "readonly") {
  const db = await openMaterialDb();
  return db.transaction(MATERIAL_STORE, mode).objectStore(MATERIAL_STORE);
}

async function sightWordStore(mode = "readonly") {
  const db = await openMaterialDb();
  return db.transaction(SIGHT_WORD_STORE, mode).objectStore(SIGHT_WORD_STORE);
}

function requestToPromise(request) {
  return new Promise((resolve, reject) => {
    request.addEventListener("success", () => resolve(request.result));
    request.addEventListener("error", () => reject(request.error));
  });
}

async function getAllMaterials() {
  return requestToPromise((await materialStore()).getAll());
}

async function getMaterial(id) {
  return requestToPromise((await materialStore()).get(id));
}

async function saveMaterial(material) {
  return requestToPromise((await materialStore("readwrite")).put(material));
}

async function deleteMaterial(id) {
  return requestToPromise((await materialStore("readwrite")).delete(id));
}

async function clearAllMaterials() {
  return requestToPromise((await materialStore("readwrite")).clear());
}

function isYoungFrankMaterial(material) {
  return /young\s+frank/i.test([
    material.title,
    material.fileName,
    material.source,
    material.notes,
    material.extractedText,
    material.weeklyPlanData?.readingTitle,
    ...(material.skillTags ?? []),
    ...(material.topicTags ?? [])
  ].join(" "));
}

function cleanWrongWeekTagText(value) {
  return String(value || "")
    .replace(/\b(?:calendar\s*week|cw|week)\s*#?\s*(?:24|26)\b/gi, "")
    .replace(/\s{2,}/g, " ")
    .replace(/\s+,/g, ",")
    .replace(/,\s*,/g, ",")
    .replace(/^,\s*|\s*,$/g, "")
    .trim();
}

function cleanWrongWeekTags(tags = []) {
  return uniqueWords(
    tags
      .map(cleanWrongWeekTagText)
      .filter((tag) => tag && !/^(?:calendar\s*week|cw|week)?\s*#?\s*(?:24|26)$/i.test(tag))
  );
}

function fixYoungFrankMaterialWeekTag(material) {
  if (!isYoungFrankMaterial(material)) return material;
  const topicTags = uniqueWords([...cleanWrongWeekTags(material.topicTags), "CW22"]);
  const skillTags = cleanWrongWeekTags(material.skillTags);
  const notes = cleanWrongWeekTagText(material.notes);
  const changed = JSON.stringify(topicTags) !== JSON.stringify(material.topicTags ?? []) ||
    JSON.stringify(skillTags) !== JSON.stringify(material.skillTags ?? []) ||
    notes !== (material.notes || "");
  return changed ? { ...material, topicTags, skillTags, notes, updatedAt: new Date().toISOString() } : material;
}

async function migrateYoungFrankCw22Tags(materials) {
  if (store.materialTagMigrationVersion === YOUNG_FRANK_CW22_TAG_MIGRATION) return materials;
  const migrated = materials.map(fixYoungFrankMaterialWeekTag);
  const changed = migrated.filter((material, index) => material !== materials[index]);
  if (changed.length) await Promise.all(changed.map(saveMaterial));
  store.materialTagMigrationVersion = YOUNG_FRANK_CW22_TAG_MIGRATION;
  saveStore();
  return migrated;
}

async function seedSightWordBank() {
  try {
    const storeObject = await sightWordStore("readwrite");
    await Promise.all(
      sightWordSeedBank.map((item) =>
        requestToPromise(storeObject.put({ ...item, word: normalizeWord(item.word), displayWord: item.word, updatedAt: new Date().toISOString() }))
      )
    );
  } catch {
    // The built-in constant still keeps the core bank available if IndexedDB is blocked.
  }
}

async function loadMaterials() {
  if (!parentIsAuthenticated()) {
    state.materials = [];
    renderMaterials();
    return;
  }
  state.materialsLoading = true;
  state.materialsStatus = "";
  renderMaterials();
  try {
    const materials = await migrateYoungFrankCw22Tags(await getAllMaterials());
    state.materials = materials.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  } catch {
    state.materialsStatus = "Could not load local materials. IndexedDB may be blocked in this browser.";
    state.materials = [];
  }
  state.materialsLoading = false;
  renderMaterials();
}

async function uploadMaterial(event) {
  event.preventDefault();
  if (!parentIsAuthenticated()) {
    showParentAuth("Please unlock Parent before uploading.");
    return;
  }

  const files = Array.from(document.querySelector("#materialFileInput").files ?? []);
  if (!files.length) {
    state.materialsStatus = "Choose at least one file first.";
    renderMaterials();
    return;
  }

  const invalid = files
    .map((file) => ({ file, error: validateMaterialFile(file) }))
    .find((item) => item.error);
  if (invalid) {
    state.materialsStatus = `${invalid.file.name}: ${invalid.error}`;
    renderMaterials();
    return;
  }

  const baseTitle = document.querySelector("#materialTitleInput").value.trim();
  const sharedFields = {
    grade: document.querySelector("#materialGradeInput").value,
    semester: document.querySelector("#materialSemesterInput").value,
    source: document.querySelector("#materialSourceInput").value.trim(),
    materialKind: document.querySelector("#materialKindInput").value,
    skillTags: parseTags(document.querySelector("#materialSkillTagsInput").value),
    topicTags: parseTags(document.querySelector("#materialTopicTagsInput").value),
    notes: document.querySelector("#materialNotesInput").value.trim()
  };
  const materials = files.map((file, index) => {
    const now = new Date().toISOString();
    const title = baseTitle
      ? files.length > 1 ? `${baseTitle} - ${index + 1}` : baseTitle
      : file.name.replace(/\.[^.]+$/, "");
    return {
      id: createId(),
      title,
      fileName: file.name,
      fileType: materialExtension(file.name),
      mimeType: file.type || "",
      fileSize: file.size,
      fileBlob: file,
      ...sharedFields,
      extractedText: "",
      detectedSkills: [],
      detectedWords: [],
      questionTypes: [],
      parseStatus: "parsing",
      parseMessage: "Parsing material...",
      createdAt: now,
      updatedAt: now
    };
  });

  state.materialsStatus = `Saving ${materials.length} material${materials.length > 1 ? "s" : ""}...`;
  await Promise.all(materials.map(saveMaterial));
  event.target.reset();
  await loadMaterials();

  let parsedCount = 0;
  const parsedMaterials = [];
  let weeklySuggestionCount = 0;
  let mainStoryCount = 0;
  for (const material of materials) {
    try {
      state.materialsStatus = `Parsing ${material.fileName}...`;
      renderMaterials();
      const extractedText = await extractMaterialText(material);
      const analysis = analyzeMaterialText(extractedText, material);
      const hasText = Boolean(extractedText.trim());
      const isPdf = material.mimeType === "application/pdf" || material.fileType === "pdf";
      const parsedWeeklyPlan = material.materialKind === "weekly_plan" && weeklyPlanDataHasContent(analysis.weeklyPlanData);
      const storyDetails = material.materialKind === "main_story_pages" ? extractMainStoryDetailsFromText(extractedText, baseTitle || material.title) : null;
      const parsedMainStory = Boolean(storyDetails?.readingArticleEvidence);
      const parseStatus = parsedMainStory
        ? "parsed_main_story"
        : parsedWeeklyPlan
        ? "parsed_weekly_plan"
        : hasText
          ? "parsed"
          : isPdf
            ? "needs_review_ocr"
            : "needs_review";
      const parseMessage = parsedMainStory
        ? "Main story pages parsed. Create or update a weekly plan to use them for Level 2."
        : parsedWeeklyPlan
        ? "Weekly plan parsed. Review suggestions before applying."
        : hasText
          ? "Parsed. Review and apply suggestions when ready."
          : isPdf
            ? "No readable PDF text found. This may be a scanned PDF; OCR or manual review is needed."
            : "No readable text found. Add tags or try OCR/source file.";
      const savedMaterial = {
        ...material,
        extractedText,
        detectedSkills: analysis.detectedSkills,
        detectedWords: analysis.detectedWords,
        questionTypes: analysis.questionTypes,
        weeklyPlanData: analysis.weeklyPlanData,
        storyDetails,
        topicTags: uniqueWords([...material.topicTags, ...analysis.topicTags]),
        parseStatus,
        parseMessage,
        updatedAt: new Date().toISOString()
      };
      await saveMaterial(savedMaterial);
      parsedMaterials.push(savedMaterial);
      if (hasText) parsedCount += 1;
      if (parsedWeeklyPlan) {
        weeklySuggestionCount += 1;
      }
      if (parsedMainStory) {
        mainStoryCount += 1;
      }
    } catch {
      const failedMaterial = {
        ...material,
        parseStatus: "failed",
        parseMessage: "Parsing failed. The file is saved; add tags manually or try another format.",
        updatedAt: new Date().toISOString()
      };
      await saveMaterial(failedMaterial);
      parsedMaterials.push(failedMaterial);
    }
  }
  await loadMaterials();
  state.weeklySuggestions = null;
  state.materialBatch = {
    id: createId(),
    title: baseTitle || `${materials.length} uploaded weekly plan file${materials.length > 1 ? "s" : ""}`,
    materialIds: parsedMaterials.map((item) => item.id),
    fileCount: materials.length,
    parsedCount,
    weeklySuggestionCount: weeklySuggestionCount + mainStoryCount,
    mainStoryCount,
    createdAt: new Date().toISOString()
  };
  state.materialsStatus = weeklySuggestionCount || mainStoryCount
    ? `${materials.length} file${materials.length > 1 ? "s" : ""} uploaded. ${weeklySuggestionCount} weekly plan page${weeklySuggestionCount === 1 ? "" : "s"} and ${mainStoryCount} main story page batch${mainStoryCount === 1 ? "" : "es"} parsed. Create one weekly plan when ready.`
    : `${materials.length} file${materials.length > 1 ? "s" : ""} uploaded. ${parsedCount} parsed; no weekly plan sections detected yet.`;
  renderMaterials();
}

async function openMaterial(id) {
  if (!parentIsAuthenticated()) {
    showParentAuth("Please unlock Parent before opening files.");
    return;
  }
  const material = await getMaterial(id);
  if (!material?.fileBlob) {
    state.materialsStatus = "Could not find the local file.";
    renderMaterials();
    return;
  }
  const url = URL.createObjectURL(material.fileBlob);
  window.open(url, "_blank", "noopener,noreferrer");
  setTimeout(() => URL.revokeObjectURL(url), 60 * 1000);
}

async function deleteMaterialAndTrainingPlans(id) {
  if (!parentIsAuthenticated()) {
    showParentAuth("Please unlock Parent before deleting files.");
    return;
  }
  const material = await getMaterial(id);
  if (!material) {
    state.materialsStatus = "Could not find this material.";
    renderMaterials();
    return;
  }
  const title = material.title || material.fileName || "this material";
  const confirmed = window.confirm(`Delete "${title}"?\n\nThis will also clear current and pending training plans generated from uploaded materials.`);
  if (!confirmed) return;
  try {
    await deleteMaterial(id);
    resetLocalTrainingPlans();
    state.weeklySuggestions = null;
    saveStore();
    await loadMaterials();
    state.materialsStatus = `${title} deleted. Training plans were reset.`;
    renderMaterials();
  } catch {
    state.materialsStatus = "Could not delete this material. IndexedDB may be blocked in this browser.";
    renderMaterials();
  }
}

function materialTextForWeekDetection(material) {
  return [
    material?.title,
    material?.fileName,
    material?.source,
    material?.notes,
    ...(material?.skillTags ?? []),
    ...(material?.topicTags ?? []),
    ...(material?.detectedSkills ?? [])
  ].join(" ");
}

function calendarWeekFromText(text) {
  const value = String(text || "");
  const match = value.match(/\b(?:calendar\s*week|cw|week)\s*#?\s*0?(\d{1,2})\b/i) ||
    value.match(/\b\d{4}-cw0?(\d{1,2})\b/i);
  return match ? Number(match[1]) : null;
}

function materialMatchesCalendarWeek(material, weekNumber) {
  if (!weekNumber) return false;
  const text = materialTextForWeekDetection(material);
  return calendarWeekFromText(text) === weekNumber;
}

function currentUploadedMaterialIdsForClear() {
  const explicitIds = uniqueWords([
    ...(state.weeklySuggestions?.sourceMaterialIds ?? []),
    ...(state.materialBatch?.materialIds ?? [])
  ]);
  if (explicitIds.length) return explicitIds;

  const weekNumber = calendarWeekFromText(state.weeklySuggestions?.sourceTitle) ||
    calendarWeekFromText(state.materialBatch?.title) ||
    nextWeekInfoForDate().calendarWeek;
  return (state.materials ?? [])
    .filter((material) => materialMatchesCalendarWeek(material, weekNumber))
    .map((material) => material.id);
}

async function clearCurrentUploadedWeeklyPlanMaterials() {
  if (!parentIsAuthenticated()) {
    showParentAuth("Please unlock Parent before clearing uploaded materials.");
    return;
  }
  const targetIds = currentUploadedMaterialIdsForClear();
  if (!targetIds.length) {
    state.materialsStatus = "No current uploaded weekly plan files found to clear.";
    renderMaterials();
    return;
  }
  const confirmed = window.confirm(`Clear ${targetIds.length} current weekly plan upload${targetIds.length === 1 ? "" : "s"}?\n\nOlder uploaded weeks will stay saved.`);
  if (!confirmed) return;
  try {
    await Promise.all(targetIds.map(deleteMaterial));
    state.materialBatch = null;
    state.weeklySuggestions = null;
    await loadMaterials();
    state.materialsStatus = "Current weekly plan uploads cleared. Older uploaded weeks were kept.";
    renderMaterials();
  } catch {
    state.materialsStatus = "Could not clear current uploads. IndexedDB may be blocked in this browser.";
    renderMaterials();
  }
}

async function extractMaterialText(material) {
  const file = material.fileBlob;
  if (material.mimeType === "application/pdf" || material.fileType === "pdf") return extractPdfText(file);
  if (["xlsx", "xls"].includes(material.fileType)) return extractSpreadsheetText(file);
  if (material.fileType === "csv" || material.mimeType === "text/csv") return file.text();
  if (material.mimeType?.startsWith("image/")) return extractImageText(file);
  return "";
}

function loadExternalScript(url, globalName) {
  if (window[globalName]) return Promise.resolve(window[globalName]);
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[data-lazy-lib="${globalName}"]`);
    if (existing) {
      existing.addEventListener("load", () => resolve(window[globalName]), { once: true });
      existing.addEventListener("error", reject, { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = url;
    script.async = true;
    script.dataset.lazyLib = globalName;
    script.onload = () => window[globalName] ? resolve(window[globalName]) : reject(new Error(`${globalName} did not initialize`));
    script.onerror = () => reject(new Error(`Could not load ${globalName}`));
    document.head.appendChild(script);
  });
}

async function loadPdfLib() {
  if (!pdfLibPromise) {
    pdfLibPromise = import(PDFJS_URL).then((module) => {
      window.pdfjsLib = module;
      module.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
      return module;
    });
  }
  return pdfLibPromise;
}

async function loadXlsxLib() {
  if (!xlsxLibPromise) xlsxLibPromise = loadExternalScript(XLSX_URL, "XLSX");
  return xlsxLibPromise;
}

async function loadTesseractLib() {
  if (!tesseractLibPromise) tesseractLibPromise = loadExternalScript(TESSERACT_URL, "Tesseract");
  return tesseractLibPromise;
}

async function extractPdfText(file) {
  const pdfjsLib = await loadPdfLib();
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  return extractTextFromPdfDocument(pdf);
}

async function extractTextFromPdfDocument(pdf) {
  const pages = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    pages.push(content.items.map((item) => item.str).join(" "));
  }
  return pages.join("\n");
}

async function extractPdfTextFromUrl(url) {
  const pdf = await loadPdfDocumentFromUrl(url);
  return extractTextFromPdfDocument(pdf);
}

function loadPdfDocumentFromUrl(url) {
  if (!pdfDocumentPromisesByUrl.has(url)) {
    const promise = loadPdfLib()
      .then((pdfjsLib) => pdfjsLib.getDocument({ url }).promise)
      .catch((error) => {
        pdfDocumentPromisesByUrl.delete(url);
        throw error;
      });
    pdfDocumentPromisesByUrl.set(url, promise);
  }
  return pdfDocumentPromisesByUrl.get(url);
}

async function renderRazPdfViewer() {
  const viewer = document.querySelector("[data-raz-pdf-viewer]");
  if (!viewer) return;
  const url = viewer.dataset.pdfUrl || "";
  const articleId = viewer.dataset.articleId || "";
  const pagesTarget = viewer.querySelector("[data-raz-pdf-pages]");
  const statusTarget = viewer.querySelector("[data-raz-pdf-status]");
  if (!url || !pagesTarget || !statusTarget) return;

  const renderToken = ++state.razPdfRenderToken;
  state.razPdfRenderError = "";
  viewer.setAttribute("aria-busy", "true");
  try {
    const pdf = await loadPdfDocumentFromUrl(url);
    if (renderToken !== state.razPdfRenderToken || !viewer.isConnected || viewer.dataset.articleId !== articleId) return;
    pagesTarget.innerHTML = "";
    const availableWidth = Math.max(280, Math.floor(viewer.clientWidth - 24));
    const outputScale = Math.min(window.devicePixelRatio || 1, 2);

    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      if (renderToken !== state.razPdfRenderToken || !viewer.isConnected) return;
      statusTarget.textContent = `Loading page ${pageNumber} of ${pdf.numPages}...`;
      const page = await pdf.getPage(pageNumber);
      const baseViewport = page.getViewport({ scale: 1 });
      const cssScale = availableWidth / baseViewport.width;
      const renderViewport = page.getViewport({ scale: cssScale * outputScale });
      const canvas = document.createElement("canvas");
      const pageFrame = document.createElement("div");
      pageFrame.className = "raz-pdf-page";
      pageFrame.setAttribute("role", "img");
      pageFrame.setAttribute("aria-label", `Page ${pageNumber} of ${pdf.numPages}`);
      canvas.width = Math.ceil(renderViewport.width);
      canvas.height = Math.ceil(renderViewport.height);
      canvas.style.width = `${Math.ceil(renderViewport.width / outputScale)}px`;
      canvas.style.height = `${Math.ceil(renderViewport.height / outputScale)}px`;
      pageFrame.appendChild(canvas);
      pagesTarget.appendChild(pageFrame);
      await page.render({ canvasContext: canvas.getContext("2d", { alpha: false }), viewport: renderViewport }).promise;
    }
    if (renderToken !== state.razPdfRenderToken || !viewer.isConnected) return;
    statusTarget.textContent = `${pdf.numPages} pages ready`;
    viewer.setAttribute("aria-busy", "false");
  } catch (error) {
    if (renderToken !== state.razPdfRenderToken || !viewer.isConnected) return;
    state.razPdfRenderError = error?.message || "The book could not be displayed.";
    viewer.setAttribute("aria-busy", "false");
    statusTarget.textContent = "The book could not be displayed here.";
    pagesTarget.innerHTML = `<p class="raz-pdf-error">Please use Full screen while the book reloads.</p>`;
  }
}

function razArticleById(articleId) {
  return razLibrary().find((article) => article.id === articleId) || null;
}

function razAssignmentsFromProgress(dayProgress = ensureAdventureDayProgress(activeProgress())) {
  return (dayProgress.razReadingAssignments ?? []).map(razArticleById).filter(Boolean);
}

function looksLikeInformationalText(text, details) {
  const flat = String(text || "").toLowerCase();
  const informationalSignals = (flat.match(/\b(?:facts?|species|scientist|animals?|plants?|water|earth|history|body|habitat|learn|means|types?|used to|can be)\b/g) ?? []).length;
  const narrativeSignals = (flat.match(/\b(?:said|asked|replied|one day|suddenly|wanted|felt|friend|mother|father)\b/g) ?? []).length;
  return informationalSignals > narrativeSignals + 2 || (!(details.readingArticleCharacters ?? []).length && !details.readingArticleProblem);
}

async function ensureRazArticleDetails(article) {
  if (!article || store.razArticleDetails?.[article.id]?.status) return store.razArticleDetails?.[article.id];
  if (state.razArticleAnalysisPending[article.id]) return state.razArticleAnalysisPending[article.id];
  state.razArticleAnalysisPending[article.id] = (async () => {
    try {
      const text = await extractPdfTextFromUrl(article.path);
      const details = extractMainStoryDetailsFromText(text, article.title);
      store.razArticleDetails = store.razArticleDetails ?? {};
      store.razArticleDetails[article.id] = {
        ...details,
        status: details.readingArticleContentStatus === "uploaded-pages" ? "ready" : "failed",
        genre: looksLikeInformationalText(text, details) ? "informational" : "literary",
        updatedAt: new Date().toISOString()
      };
      saveStore();
    } catch (error) {
      store.razArticleDetails = store.razArticleDetails ?? {};
      store.razArticleDetails[article.id] = {
        status: "failed",
        genre: "unknown",
        error: error?.message || "PDF text could not be read.",
        updatedAt: new Date().toISOString()
      };
      saveStore();
    } finally {
      delete state.razArticleAnalysisPending[article.id];
      if (currentPracticeRoute()?.practiceId === "read-warmup" || currentPracticeRoute()?.practiceId === "speak") renderPracticeRoute();
    }
    return store.razArticleDetails?.[article.id];
  })();
  return state.razArticleAnalysisPending[article.id];
}

function queueRazArticleAnalysis(articles = []) {
  articles.forEach((article) => {
    if (!store.razArticleDetails?.[article.id]?.status && !state.razArticleAnalysisPending[article.id]) {
      ensureRazArticleDetails(article);
    }
  });
}

async function extractSpreadsheetText(file) {
  const XLSX = await loadXlsxLib();
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
  return workbook.SheetNames.map((sheetName) => {
    const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
    return `${sheetName}\n${csv}`;
  }).join("\n\n");
}

function englishOnlyOcrText(text) {
  return String(text || "")
    .replace(/[^\x00-\x7F]+/g, " ")
    .split(/\n+/)
    .map((line) => line
      .replace(/[•·●]/g, "\n")
      .replace(/[^\w'",:;./()&\-\s]+/g, " ")
      .replace(/\s+/g, " ")
      .trim()
    )
    .filter(Boolean)
    .join("\n");
}

async function extractImageText(file) {
  const Tesseract = await loadTesseractLib();
  const result = await Tesseract.recognize(file, "eng");
  const rawText = result.data.text || "";
  const rightColumnText = rightColumnTextFromOcrResult(result);
  const englishOnlyText = englishOnlyOcrText(rawText);
  const englishOnlyRightColumnText = englishOnlyOcrText(rightColumnText);
  const sourcePageType = weeklyPlanPageType(`${rawText}\n${rightColumnText}`);
  return [
    rawText,
    `ENGLISH_ONLY_TEXT:\n${englishOnlyText}\nEND_ENGLISH_ONLY_TEXT`,
    `OCR_PARSER_VERSION: ${WEEKLY_PLAN_OCR_VERSION}`,
    `SOURCE_PAGE_TYPE: ${sourcePageType}`,
    englishOnlyRightColumnText ? `RIGHT_COLUMN_WORDS:\n${englishOnlyRightColumnText}\nEND_RIGHT_COLUMN_WORDS` : ""
  ].filter(Boolean).join("\n\n");
}

function ocrWordBox(word = {}) {
  const box = word.bbox || word;
  const x0 = Number(box.x0 ?? box.left ?? box.x ?? 0);
  const y0 = Number(box.y0 ?? box.top ?? box.y ?? 0);
  const x1 = Number(box.x1 ?? (box.left != null && box.width != null ? box.left + box.width : 0));
  const y1 = Number(box.y1 ?? (box.top != null && box.height != null ? box.top + box.height : 0));
  if (!Number.isFinite(x0) || !Number.isFinite(y0) || !Number.isFinite(x1) || !Number.isFinite(y1)) return null;
  if (x1 <= x0 || y1 <= y0) return null;
  return { x0, y0, x1, y1, cx: (x0 + x1) / 2, cy: (y0 + y1) / 2, height: y1 - y0 };
}

function tesseractWords(result = {}) {
  const words = result.data?.words ?? [];
  return words
    .map((word) => ({ text: String(word.text || "").trim(), box: ocrWordBox(word) }))
    .filter((word) => word.text && word.box);
}

function groupOcrWordsIntoLines(words = []) {
  return groupOcrWordObjectsIntoLines(words)
    .map((line) => line.text)
    .filter(Boolean);
}

function groupOcrWordObjectsIntoLines(words = []) {
  const sorted = [...words].sort((a, b) => a.box.cy - b.box.cy || a.box.x0 - b.box.x0);
  const lines = [];
  sorted.forEach((word) => {
    const tolerance = Math.max(8, word.box.height * 0.75);
    const line = lines.find((item) => Math.abs(item.cy - word.box.cy) <= tolerance);
    if (line) {
      line.words.push(word);
      line.cy = (line.cy * (line.words.length - 1) + word.box.cy) / line.words.length;
    } else {
      lines.push({ cy: word.box.cy, words: [word] });
    }
  });
  return lines
    .map((line) => {
      const lineWords = line.words.sort((a, b) => a.box.x0 - b.box.x0);
      return {
        cy: line.cy,
        x0: Math.min(...lineWords.map((word) => word.box.x0)),
        x1: Math.max(...lineWords.map((word) => word.box.x1)),
        words: lineWords,
        text: lineWords.map((word) => word.text).join(" ").replace(/\s+/g, " ").trim()
      };
    })
    .filter((line) => line.text);
}

function rightColumnTextFromOcrResult(result = {}) {
  const words = tesseractWords(result);
  if (!words.length) return "";
  const maxX = Math.max(...words.map((word) => word.box.x1));
  const minX = Math.min(...words.map((word) => word.box.x0));
  const width = Math.max(1, maxX - minX);
  const lines = groupOcrWordObjectsIntoLines(words);
  const headerLine = lines.find((line) => /bilingual/i.test(line.text) && /vocabulary/i.test(line.text)) ||
    lines.find((line) => /vocabulary/i.test(line.text) && line.x0 > minX + width * 0.65);
  const threshold = headerLine ? Math.max(minX + width * 0.70, headerLine.x0 - 16) : minX + width * 0.82;
  const minY = headerLine ? headerLine.cy : Math.min(...words.map((word) => word.box.y0));
  const rightWords = words
    .filter((word) => word.box.cx >= threshold)
    .filter((word) => word.box.cy >= minY)
    .filter((word) => /[A-Za-z]/.test(word.text))
    .filter((word) => !/^[^A-Za-z]+$/.test(word.text));
  return groupOcrWordsIntoLines(rightWords).join("\n");
}

const skillDetectors = [
  ["short vowels", /\bshort vowel|cvc\b|\bcat\b|\bpen\b|\bpig\b|\bdog\b|\bsun\b/i],
  ["long vowels", /\blong vowel|silent e|magic e|final-e|a_e|i_e|o_e|u_e/i],
  ["vowel teams", /\bvowel team|ai|ay|ee|ea|oa|ow|igh/i],
  ["r-controlled vowels", /\br-controlled|bossy r|ar\b|er\b|ir\b|or\b|ur\b/i],
  ["syllables", /\bsyllable|two-syllable|split the word|word part/i],
  ["prefixes", /\bprefix|un-|re-|pre-/i],
  ["suffixes", /\bsuffix|-ed|-ing|-er|-est/i],
  ["contractions", /\bcontraction|can't|didn't|isn't|that's|we'll|you'll|i'm|i'll/i],
  ["sight words", /\bsight word|high.frequency|tricky word|irregular word/i],
  ["pronouns", /\bpronoun|he\b|she\b|they\b|we\b|it\b/i],
  ["verb tense", /\bverb tense|past tense|present tense|future tense|was|were|has|have/i],
  ["punctuation", /\bpunctuation|period|question mark|comma|quotation mark|end mark/i],
  ["complete sentences", /\bcomplete sentence|sentence answer|capital letter|end mark/i],
  ["main idea", /\bmain idea|mostly about|central idea/i],
  ["details", /\bdetail|text evidence|prove your answer|support/i],
  ["sequence", /\bsequence|first|next|then|last|beginning|middle|end/i],
  ["cause and effect", /\bcause|effect|because|why did|what happened/i],
  ["retell", /\bretell|summarize|beginning middle end/i],
  ["writing response", /\bwrite|writing|short answer|respond|sentence/i]
];

const questionDetectors = [
  ["multiple choice", /\bA\.|\bB\.|\bC\.|choose the best|circle the answer/i],
  ["spelling", /\bspell|dictation|write the word/i],
  ["matching", /\bmatch|draw a line/i],
  ["fill in the blank", /\bfill in|blank|__/i],
  ["short response", /\bwrite a sentence|answer in a complete sentence|explain/i],
  ["reading comprehension", /\bread the passage|answer the question|main idea|detail/i]
];

const highValueWords = new Set([
  ...Object.keys(wordHelp),
  ...defaultWeeklyPlan.masteryWords.map(normalizeWord),
  ...defaultWeeklyPlan.sightWords.map(normalizeWord),
  "because",
  "before",
  "after",
  "first",
  "next",
  "then",
  "finally",
  "different",
  "important",
  "machine",
  "architect",
  "building",
  "design",
  "sentence",
  "question"
]);

function analyzeMaterialText(text, material) {
  const normalized = text.replace(/\s+/g, " ").trim();
  const weeklyPlanData = parseWeeklyPlanData(text);
  const detectedSkills = uniqueWords([
    ...material.skillTags,
    ...skillDetectors.filter(([, pattern]) => pattern.test(normalized)).map(([skill]) => skill),
    ...weeklyPlanData.foundationalSkills,
    ...weeklyPlanData.readingComprehension,
    ...weeklyPlanData.grammar
  ]);
  const questionTypes = uniqueWords(questionDetectors.filter(([, pattern]) => pattern.test(normalized)).map(([type]) => type));
  const words = extractCandidateWords(normalized);
  const topicTags = extractTopicTags(normalized, material);
  const structuredWords = cleanOcrVocabularyList([
    ...weeklyPlanData.vocabulary,
    ...weeklyPlanData.uoiVocabulary,
    ...weeklyPlanData.uiVocabulary,
    ...weeklyPlanData.englishVocabulary,
    ...weeklyPlanData.phonicsVocabulary,
    ...weeklyPlanData.highFrequencyWords,
    ...weeklyPlanData.spellingBasicWords,
    ...weeklyPlanData.spellingReviewWords,
    ...weeklyPlanData.spellingChallengeWords
  ]);
  const detectedWords = uniqueWords([
    ...structuredWords,
    ...(structuredWords.length ? [] : words)
  ]);
  return { detectedSkills, detectedWords, questionTypes, topicTags, weeklyPlanData };
}

function cleanedWeeklyPlanData(data = {}) {
  const englishWords = cleanEnglishVocabularyList(data.englishVocabulary ?? []);
  const englishKeys = new Set(englishWords.map(normalizeWord));
  const uiWords = cleanOcrVocabularyList(data.uiVocabulary ?? []).filter((word) => !englishKeys.has(normalizeWord(word)));
  return {
    ...data,
    vocabulary: cleanOcrVocabularyList(data.vocabulary ?? []),
    uiVocabulary: uiWords,
    englishVocabulary: englishWords,
    uoiVocabulary: cleanOcrVocabularyList(data.uoiVocabulary ?? []),
    readingSelections: uniqueWords(data.readingSelections ?? []),
    homeLearningSuggestions: uniqueWords(data.homeLearningSuggestions ?? []),
    weeklyLearningOutcomes: uniqueWords(data.weeklyLearningOutcomes ?? []),
    pageTextSignals: uniqueWords(data.pageTextSignals ?? [])
  };
}

function normalizeOcrText(text) {
  return String(text || "")
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, "\"")
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim()
    .replace(/\bHigh Frequency Word[s]?:/gi, "High Frequency Words:")
    .replace(/\bBASIC WORD[s]?\b\s*:*/gi, "BASIC WORDS:")
    .replace(/\bREVIEW\s*:/g, "REVIEW:")
    .replace(/\bCHALLENGE\s*:/g, "CHALLENGE:");
}

function flattenOcrText(text) {
  return normalizeOcrText(text)
    .replace(/\s+/g, " ")
    .trim();
}

function structuredOcrMarker(text, marker) {
  const match = String(text || "").match(new RegExp(`(?:^|\\n)\\s*${marker}\\s*:\\s*([^\\n]+)`, "i"));
  return match?.[1]?.trim() || "";
}

function structuredOcrBlock(text, marker, endMarker) {
  const pattern = new RegExp(`(?:^|\\n)\\s*${marker}\\s*:\\s*\\n?([\\s\\S]*?)(?:\\n\\s*${endMarker}\\b|$)`, "i");
  return String(text || "").match(pattern)?.[1]?.trim() || "";
}

function weeklyPlanEnglishText(text) {
  const normalized = normalizeOcrText(text);
  const englishOnly = structuredOcrBlock(normalized, "ENGLISH_ONLY_TEXT", "END_ENGLISH_ONLY_TEXT");
  return normalizeOcrText(englishOnly || englishOnlyOcrText(text));
}

function removeStructuredOcrBlocks(text) {
  return String(text || "")
    .replace(/(?:^|\n)\s*ENGLISH_ONLY_TEXT\s*:\s*[\s\S]*?(?:\n\s*END_ENGLISH_ONLY_TEXT\b|$)/gi, "\n")
    .replace(/(?:^|\n)\s*RIGHT_COLUMN_WORDS\s*:\s*[\s\S]*?(?:\n\s*END_RIGHT_COLUMN_WORDS\b|$)/gi, "\n")
    .replace(/(?:^|\n)\s*OCR_PARSER_VERSION\s*:[^\n]*/gi, "\n")
    .replace(/(?:^|\n)\s*SOURCE_PAGE_TYPE\s*:[^\n]*/gi, "\n");
}

function englishSubjectBodyText(text) {
  return flattenOcrText(removeStructuredOcrBlocks(weeklyPlanEnglishText(text)));
}

function wordsFromStrictList(text, { limit = 16, keepHighFrequency = true } = {}) {
  return uniqueWords(
    String(text || "")
      .replace(/\band\b/gi, ",")
      .split(/,|;|\n|\s+/)
      .map((item) => item.trim().replace(/^[^A-Za-z']+|[^A-Za-z']+$/g, ""))
      .filter((item) => /^[A-Za-z][A-Za-z']*$/.test(item))
      .map(normalizeWord)
      .filter((word) => word.length > 1)
      .filter((word) => !/^(?:basic|words?|review|challenge|challenging|decodable|skill|skills|diphthongs|spelling|bilingual|vocabulary)$/i.test(word))
      .filter((word) => keepHighFrequency || !likelyHighFrequencyWords.has(word))
      .filter((word) => !spellingOcrFalseWords.has(word))
      .filter((word) => !/(?:ii|ff|kk|jj)/.test(word))
  ).slice(0, limit);
}

function bilingualEntriesFromRightColumn(text) {
  const blocked = new Set([
    "bilingual", "vocabulary", "subject", "students", "will", "learning", "outcomes", "english", "uoi",
    "feat", "rhe", "ba", "hl", "ard", "eo", "ew", "jump", "child", "read", "discuss", "learn",
    "review", "create", "ask", "home", "learning", "suggestions", "student", "students"
  ]);
  return uniqueWords(
    String(text || "")
      .split(/\n+/)
      .map((line) => line.trim().replace(/^[^A-Za-z']+/, ""))
      .map((line) => {
        const words = line.match(/[A-Za-z][A-Za-z']*/g) ?? [];
        return bestBilingualWordFromLine(words);
      })
      .filter(Boolean)
      .filter((word) => !blocked.has(word))
  );
}

function bilingualEntriesFromBilingualSection(text) {
  const source = normalizeOcrText(text);
  const start = source.search(/\bBilingual Vocabulary\b/i);
  if (start < 0) return [];
  const after = source.slice(start).replace(/\bBilingual Vocabulary\b[^\n]*/i, "");
  const section = after.split(/\b(?:Home Learning Suggestions|Students will|Learning Outcomes|SOURCE_PAGE_TYPE|OCR_PARSER_VERSION)\b/i)[0] || after;
  const lineEntries = bilingualEntriesFromRightColumn(section);
  return lineEntries.length >= 3 ? lineEntries : [];
}

function strictBilingualToken(word) {
  const normalized = normalizeWeeklyOcrWord(word);
  if (!normalized || normalized.length < 2) return "";
  if (/^(?:feat|rhe|ba|hl|ard|eo|ew|ex|hh|fii|ufflk|iif|hii|jiil|mod|wk)$/i.test(normalized)) return "";
  if (/^[a-z]{2}$/.test(normalized) && !["ow", "ou"].includes(normalized)) return "";
  if (/(?:ii|ff|kk|jj)/.test(normalized)) return "";
  if (normalized.length > 14) return "";
  return normalized;
}

function bestBilingualWordFromLine(words = []) {
  const candidates = words
    .map(strictBilingualToken)
    .filter(Boolean)
    .filter((word) => !/^(?:bilingual|vocabulary|subject|students|will|learning|outcomes|english|uoi)$/i.test(word));
  if (!candidates.length) return "";
  return candidates[0];
}

function trustedBilingualEntries(rightColumnText, flatText, sourcePageType, fullText = "") {
  const entries = uniqueWords([
    ...bilingualEntriesFromRightColumn(rightColumnText),
    ...bilingualEntriesFromBilingualSection(fullText)
  ]);
  const max = sourcePageType === "englishSubject" ? 18 : 10;
  return entries
    .filter((word) => word.length >= 3 || ["ow", "ou"].includes(word))
    .filter((word) => !spellingOcrFalseWords.has(word))
    .filter((word) => !/^(?:vol|fuk|filifi|co|rhe)$/i.test(word))
    .slice(0, max);
}

function wordsFromSection(text) {
  return uniqueWords(
    String(text || "")
      .replace(/\band\b/gi, ",")
      .split(/,|;|\n|•|·|\s+/)
      .map((item) => item.trim().replace(/^[^A-Za-z']+|[^A-Za-z']+$/g, ""))
      .filter((item) => /^[A-Za-z][A-Za-z']*$/.test(item))
      .filter((item) => !/^(?:decodable|skill|skills|diphthongs|basic|words?|review|challenge|spelling|jump|rope|reader|writing|workshop|grammar|day|activity)$/i.test(item))
  );
}

const spellingOcrNoiseWords = new Set([
  "and", "the", "all", "they", "make", "work", "easier", "by", "asking", "what", "is", "problem",
  "tool", "can", "help", "reflect", "we", "use", "it", "safely", "celebrate", "choose", "machine",
  "to", "then", "explain", "if", "helps", "favourite", "favorite", "fit", "on", "learned", "share",
  "their", "understand", "humans", "have", "changed", "world", "using", "tools", "write", "read",
  "where", "syllable", "counting", "phonemes", "blends", "ow", "ou", "ii", "six", "simple",
  "machines", "wheel", "pulley", "lever", "screw", "wedge", "ramp", "you"
]);

const spellingOcrFalseWords = new Set([
  "ab", "vone", "hove", "heord", "wehand", "teffff", "gownn", "ral", "tey", "hore",
  "jiil", "ufflk", "hii", "iif", "rik", "jz", "cli", "kity", "iff"
]);

const likelyHighFrequencyWords = new Set([
  "a", "all", "am", "an", "and", "are", "as", "at", "away", "be", "because", "been", "before",
  "big", "boy", "by", "can", "come", "could", "did", "do", "down", "find", "first", "for",
  "found", "from", "go", "good", "had", "has", "have", "he", "her", "here", "him", "his",
  "how", "i", "in", "is", "it", "like", "little", "look", "make", "me", "my", "new",
  "no", "not", "now", "of", "on", "one", "or", "out", "play", "said", "see", "she", "so",
  "some", "that", "the", "their", "then", "there", "they", "this", "to", "too", "up", "use",
  "was", "we", "what", "where", "with", "you"
]);

function spellingWordsFromSection(text, limit = 14) {
  return uniqueWords(
    String(text || "")
      .replace(/\b(?:BASIC WORDS?|REVIEW|CHALLENGE|Spelling)\s*:?\b/gi, " ")
      .replace(/[•·]/g, " ")
      .split(/,|;|\n|\s+/)
      .map((item) => item.trim().replace(/^[^A-Za-z']+|[^A-Za-z']+$/g, ""))
      .filter((item) => /^[A-Za-z][A-Za-z']*$/.test(item))
      .filter((item) => !/^[A-Z]{2,}$/.test(item))
      .filter((item) => !/[A-Z]{2,}[a-z]|[a-z][A-Z]{2,}|[A-Za-z]{14,}/.test(item))
      .filter((item) => !/^(?:decodable|skill|skills|diphthongs|basic|words?|review|challenge|spelling|jump|rope|reader|writing|workshop|grammar|day|activity)$/i.test(item))
      .map(normalizeWord)
      .filter((item) => item.length > 1)
      .filter((item) => !spellingOcrNoiseWords.has(item))
      .filter((item) => !spellingOcrFalseWords.has(item))
      .filter((item) => !/(?:ii|ff|kk|jj)/.test(item))
  ).slice(0, limit);
}

function highFrequencyWordsFromSection(text, flatText) {
  const source = `${text || ""}\n${flatText || ""}`;
  const decodableMatch = String(source).match(/\bDecodable\s*:?\s*([\s\S]*?)(?:\b(?:BASIC WORDS?|REVIEW|CHALLENGE|SKILL|Reading Workshop|Writing Workshop|Jump Rope|Home Learning|SOURCE_PAGE_TYPE|RIGHT_COLUMN_WORDS)\b|[.;]\s+[A-Z][a-z]+|$)/i);
  const decodableWords = decodableMatch ? wordsFromStrictList(decodableMatch[1], { limit: 12 }) : [];
  const decodableHighFrequency = decodableWords.filter((word) => likelyHighFrequencyWords.has(word));
  if (decodableHighFrequency.length) return uniqueWords(decodableHighFrequency).slice(0, 10);

  if (!String(text || "").trim()) return [];
  const directWords = wordsFromSection(text)
    .map(normalizeWord)
    .filter((word) => likelyHighFrequencyWords.has(word))
    .filter((word) => !spellingOcrNoiseWords.has(word) || ["boy", "down", "found", "how", "now", "out"].includes(word));
  if (directWords.length) return uniqueWords(directWords).slice(0, 10);

  const fallbackBank = [
    "boy", "down", "found", "how", "now", "out", "because", "before", "after", "first", "next", "then",
    "been", "heard", "hurry", "learn", "loved", "often", "study", "world"
  ];
  return knownWordsFromText(flatText, fallbackBank).slice(0, 10);
}

function englishSubjectVocabularyFallback(text) {
  const flat = flattenOcrText(text);
  const highFrequencyMatch = flat.match(/\bHigh Frequency Words\s*:?\s*([\s\S]*?)(?:\bDecodable\b|\bSpelling\b|\bBASIC WORDS?\b|\bREVIEW\b|\bCHALLENGE\b|$)/i)?.[1] || "";
  const decodableMatch = flat.match(/\bDecodable\s*:?\s*([\s\S]*?)(?:\bBilingual Vocabulary\b|\bSpelling\b|\bBASIC WORDS?\b|\bREVIEW\b|\bCHALLENGE\b|\breview syllable\b|\bread the jump\b|\bHome Learning\b|$)/i)?.[1] || "";
  return uniqueWords([
    ...wordsFromStrictList(highFrequencyMatch, { limit: 12 }),
    ...wordsFromStrictList(decodableMatch, { limit: 12 })
  ]).slice(0, 18);
}

function uoiSubjectVocabularyFallback(text) {
  const bank = knownUoiVocabulary.map(normalizeWeeklyOcrWord);
  return knownWordsFromText(text, bank)
    .map(normalizeWeeklyOcrWord)
    .filter((word) => !["wheel", "pulley", "lever", "screw", "wedge", "ramp", "machine", "machines", "safe", "safely", "easier"].includes(normalizeWord(word)))
    .slice(0, 10);
}

function repairUoiBilingualOrder(rightColumnEntries = [], fallbackEntries = []) {
  if (!rightColumnEntries.length || fallbackEntries.length < 4) return fallbackEntries.length ? fallbackEntries : rightColumnEntries;
  const fallbackSet = new Set(fallbackEntries.map(normalizeWord));
  const used = new Set();
  const missing = fallbackEntries.filter((word) => !rightColumnEntries.map(normalizeWord).includes(normalizeWord(word)));
  const repaired = rightColumnEntries.map((word) => {
    const key = normalizeWord(word);
    if (fallbackSet.has(key) && !used.has(key)) {
      used.add(key);
      return word;
    }
    const replacement = missing.find((item) => !used.has(normalizeWord(item)));
    if (replacement) {
      used.add(normalizeWord(replacement));
      return replacement;
    }
    return "";
  }).filter(Boolean);
  return uniqueWords([...repaired, ...fallbackEntries.filter((word) => !used.has(normalizeWord(word)))]);
}

function cleanSpellingGroupWords(words = [], {
  limit = 8,
  blocked = new Set(),
  allowHighFrequency = true
} = {}) {
  const knownUoi = new Set(knownUoiVocabulary.map(normalizeWord));
  return uniqueWords(words)
    .map(normalizeWord)
    .filter((word) => word.length > 1)
    .filter((word) => !blocked.has(word))
    .filter((word) => !knownUoi.has(word))
    .filter((word) => !spellingOcrNoiseWords.has(word))
    .filter((word) => !spellingOcrFalseWords.has(word))
    .filter((word) => allowHighFrequency || !likelyHighFrequencyWords.has(word))
    .filter((word) => !/(?:ii|ff|kk|jj)/.test(word))
    .slice(0, limit);
}

function normalizeWeeklyOcrWord(word) {
  const ocrVocabularyCorrections = {
    favourite: "favorite",
    refect: "reflect"
  };
  const normalized = normalizeWord(word);
  return ocrVocabularyCorrections[normalized] || normalized;
}

function cleanOcrVocabularyItem(item) {
  const text = String(item || "")
    .replace(/[^\w'\-\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!text || /\d/.test(text)) return "";
  const words = text.split(/\s+/).filter(Boolean);
  if (!words.length || words.length > 3) return "";
  const normalizedWords = words.map(normalizeWeeklyOcrWord).filter(Boolean);
  if (!normalizedWords.length) return "";
  const normalizedPhrase = normalizedWords.join(" ");
  const shortAllCapsTokens = words.filter((word) => /^[A-Z]{2,}$/.test(word));
  const sentenceStopWords = new Set([
    "a", "an", "and", "the", "their", "share", "what", "which", "how", "can",
    "will", "would", "could", "should", "is", "are", "was", "were", "to", "then", "if", "with"
  ]);
  const labelWords = new Set([
    "vocabulary", "uoi", "topic", "bilingual", "english", "power", "oral", "words", "word",
    "spelling", "basic", "review", "challenge", "skill", "skills", "decodable"
  ]);
  const stopWordCount = normalizedWords.filter((word) => sentenceStopWords.has(word)).length;
  if (shortAllCapsTokens.length) return "";
  if (normalizedWords.some((word) => labelWords.has(word))) return "";
  if (stopWordCount) return "";
  if (words.length >= 3) return "";
  if (normalizedWords.some((word) => word.length < 2 && word !== "a" && word !== "i")) return "";
  return normalizedPhrase;
}

function cleanOcrVocabularyList(items = []) {
  return uniqueWords(
    items
      .map(cleanOcrVocabularyItem)
      .filter(Boolean)
  );
}

function cleanEnglishVocabularyList(items = []) {
  const blocked = new Set(["english", "vocabulary", "oral", "power", "words", "word", ...spellingOcrFalseWords]);
  return uniqueWords(
    items
      .map((item) => String(item || "").trim())
      .flatMap((item) => item.split(/,|;|\n|•|·/))
      .map((item) => item.trim().replace(/^[^A-Za-z']+|[^A-Za-z']+$/g, "").replace(/\s+/g, " "))
      .filter(Boolean)
      .flatMap((item) => item.includes(" ") ? item.split(/\s+/) : [item])
      .map(normalizeWeeklyOcrWord)
      .filter((word) => word.length > 1)
      .filter((word) => !blocked.has(word))
      .filter((word) => !/^[A-Z]{2,}$/.test(word))
      .filter((word) => !/(?:ii|ff|kk|jj)/.test(word))
  );
}

function phrasesFromSection(text) {
  return cleanOcrVocabularyList(
    String(text || "")
      .replace(/\b(?:ORAL\s+POWER|POWER|Vocabulary|UOI|Bilingual Vocabulary|Topic Vocabulary)\s*[-:]?/gi, "")
      .replace(/\band\b/gi, ",")
      .split(/,|;|\n|•|·/)
      .map((item) => item.trim().replace(/^[^A-Za-z']+|[^A-Za-z']+$/g, "").replace(/\s+/g, " "))
      .filter((item) => /[A-Za-z]/.test(item))
      .filter((item) => item.split(/\s+/).length <= 3)
  );
}

function wordsFromVocabularyText(text) {
  return phrasesFromSection(
    String(text || "")
      .replace(/\bEnglish Vocabulary\s*:?/gi, "")
      .replace(/\bORAL\s+POWER\s*[-:]?/gi, "")
      .replace(/\bPOWER\s*[-:]?/gi, "")
      .replace(/\bVocabulary\s*:?/gi, "")
  );
}

function wordsFromEnglishVocabularyText(text) {
  const blocked = new Set([
    "english", "vocabulary", "oral", "power", "words", "word",
    ...spellingOcrFalseWords
  ]);
  return uniqueWords(
    String(text || "")
      .replace(/\bEnglish Vocabulary\s*:?/gi, " ")
      .replace(/\bORAL\s+POWER\s*[-:]?/gi, " ")
      .replace(/\bPOWER\s*[-:]?/gi, " ")
      .replace(/\bVocabulary\s*:?/gi, " ")
      .split(/,|;|\n|•|·|\s+/)
      .map((item) => item.trim().replace(/^[^A-Za-z']+|[^A-Za-z']+$/g, ""))
      .filter((item) => /^[A-Za-z][A-Za-z']*$/.test(item))
      .filter((item) => !/^[A-Z]{2,}$/.test(item))
      .map(normalizeWeeklyOcrWord)
  )
    .filter((word) => word && !blocked.has(word))
    .filter((word) => word.length > 1)
    .filter((word) => !/(?:ii|ff|kk|jj)/.test(word));
}

function isContractionWord(word) {
  return /'/.test(word) || ["isn't", "that's", "we'll", "you'll", "can't", "didn't", "i'll", "i'm", "wasn't", "it's"].includes(normalizeWord(word));
}

function knownWordsFromText(text, words) {
  const normalized = ` ${String(text || "").toLowerCase().replace(/[^a-z'\s-]/g, " ")} `;
  return uniqueWords(
    words.filter((word) => {
      const key = normalizeWord(word);
      const variants = uniqueWords([key, key === "favorite" ? "favourite" : "", key === "favourite" ? "favorite" : ""]);
      return variants.some((variant) => new RegExp(`(^|[^a-z'])${variant.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}([^a-z']|$)`, "i").test(normalized));
    })
  );
}

function applyWeeklyPlanOcrCorrections(data) {
  const correctWords = (words = []) => uniqueWords(words.map(normalizeWeeklyOcrWord).filter(Boolean));
  const spellingBasicWords = correctWords(data.spellingBasicWords ?? []);
  const spellingReviewWords = correctWords(data.spellingReviewWords ?? []);
  const spellingChallengeWords = correctWords(data.spellingChallengeWords ?? []);
  return {
    ...data,
    uiVocabulary: correctWords(data.uiVocabulary ?? []),
    uoiVocabulary: correctWords(data.uoiVocabulary ?? []),
    englishVocabulary: correctWords(data.englishVocabulary ?? []),
    highFrequencyWords: correctWords(data.highFrequencyWords ?? []),
    readingSelections: uniqueWords(data.readingSelections ?? []),
    homeLearningSuggestions: uniqueWords(data.homeLearningSuggestions ?? []),
    weeklyLearningOutcomes: uniqueWords(data.weeklyLearningOutcomes ?? []),
    pageTextSignals: uniqueWords(data.pageTextSignals ?? []),
    spelling: {
      ...(data.spelling ?? {}),
      basic: spellingBasicWords,
      review: spellingReviewWords,
      challenge: spellingChallengeWords
    },
    spellingBasicWords,
    spellingReviewWords,
    spellingChallengeWords
  };
}

function extractBetween(text, startPattern, endPatterns = []) {
  const start = text.search(startPattern);
  if (start < 0) return "";
  const after = text.slice(start).replace(startPattern, "");
  const endIndexes = endPatterns.map((pattern) => after.search(pattern)).filter((index) => index >= 0);
  const end = endIndexes.length ? Math.min(...endIndexes) : after.length;
  return after.slice(0, end).trim();
}

function extractEnglishSubjectSpellingSections(text) {
  const flat = englishSubjectBodyText(text);
  const basicStart = flat.search(/\bbasic words?\b/i);
  const spellingWindow = basicStart >= 0 ? flat.slice(basicStart, basicStart + 700) : flat;
  const hardStop = /\breview syllable\b|\breview syllable counting\b|\bread the jump\b|\bgo through\b|\bhave a friday\b|\bhome learning\b|\bask your child\b/i;
  const bounded = spellingWindow.split(hardStop)[0] || spellingWindow;
  const trimSentence = (value) => String(value || "").split(/\.(?:\s|$)|;(?:\s|$)/)[0] || "";
  const basic = bounded.match(/\bbasic words?\b\s*[:,]?\s*([\s\S]*?)(?:\b(?:review words?|review)\b\s*[:,]?|\b(?:challenge|challenging) words?\b\s*[:,]?|$)/i)?.[1] || "";
  const review = bounded.match(/\breview words?\b\s*[:,]?\s*([\s\S]*?)(?:\b(?:challenge|challenging) words?\b\s*[:,]?|$)/i)?.[1] ||
    bounded.match(/\breview\b\s*[:,]?\s*([\s\S]*?)(?:\b(?:challenge|challenging) words?\b\s*[:,]?|$)/i)?.[1] || "";
  const challenge = trimSentence(bounded.match(/\b(?:challenge|challenging) words?\b\s*[:,]?\s*([\s\S]*?)(?:$)/i)?.[1] || "");
  return { basic, review, challenge };
}

const weeklySectionMarkers = {
  vocabulary: {
    pattern: /\bVocabulary\s*:?\s*/gi,
    valid: (text, match) => !/(?:Bilingual|Topic|Generative|English)\s+$/i.test(text.slice(Math.max(0, match.index - 25), match.index))
  },
  uoi: { pattern: /\b(?:UOI|Bilingual Vocabulary|Topic Vocabulary)\s*:?\s*/gi },
  englishVocabulary: { pattern: /\bEnglish Vocabulary\s*:?\s*/gi },
  englishPower: { pattern: /(?:^|\n)\s*(?:Oral Power|Power Words|POWER)\s*:?\s*/gi },
  highFrequency: { pattern: /\bHigh Frequency Words\s*:?\s*/gi },
  spelling: { pattern: /\bSpelling\s*:?\s*/gi },
  basicWords: { pattern: /\bBASIC WORDS?\s*:?\s*/gi },
  review: { pattern: /\bREVIEW\s*:?\s*/gi },
  challenge: { pattern: /\bCHALLENGE\s*:?\s*/gi },
  foundational: { pattern: /\b(?:Foundational Skills|SKILL|SKILLS)\s*:?\s*/gi },
  comprehension: { pattern: /\bReading Workshop\s*\/?\s*Comprehension\s*:?\s*/gi },
  grammar: { pattern: /\bWriting Workshop\s*\/?\s*Grammar\s*:?\s*/gi },
  jumpRope: { pattern: /\bJump Rope Reader\s*:?\s*/gi },
  homeLearning: { pattern: /\bHome Learning\s*:?\s*/gi },
  generative: { pattern: /\bGenerative Vocabulary\s*:?\s*/gi },
  dayActivity: { pattern: /\bDay\s+Activity\s*:?\s*/gi },
  included: { pattern: /\bIncluded\s*:?\s*/gi }
};

function weeklyMarkerMatches(text, key) {
  const marker = weeklySectionMarkers[key];
  if (!marker) return [];
  const pattern = new RegExp(marker.pattern.source, marker.pattern.flags.includes("g") ? marker.pattern.flags : `${marker.pattern.flags}g`);
  const matches = [];
  let match;
  while ((match = pattern.exec(text)) !== null) {
    const candidate = { index: match.index, length: match[0].length, text: match[0] };
    if (!marker.valid || marker.valid(text, candidate)) matches.push(candidate);
    if (match[0].length === 0) pattern.lastIndex += 1;
  }
  return matches;
}

function extractWeeklySection(text, startKeys = [], endKeys = Object.keys(weeklySectionMarkers)) {
  const starts = startKeys
    .flatMap((key) => weeklyMarkerMatches(text, key))
    .sort((a, b) => a.index - b.index);
  if (!starts.length) return "";
  const start = starts[0].index + starts[0].length;
  const endIndexes = endKeys
    .flatMap((key) => weeklyMarkerMatches(text, key))
    .map((match) => match.index)
    .filter((index) => index >= start);
  const end = endIndexes.length ? Math.min(...endIndexes) : text.length;
  return text.slice(start, end).trim();
}

function parseReadingTitleFromWeeklyText(text) {
  const candidates = [
    text.match(/\b(?:Reading|Read|Story|Selection|Anchor Text|Weekly Title|Text)\s*(?:title)?\s*[:\-]\s*["']?([^"'\n.]{2,80})/i)?.[1],
    text.match(/\bread(?:ing)?\s+["']([^"']{2,80})["']/i)?.[1],
    text.match(/\bread and discuss the story\s*['"]?([^'".]{2,80})['"]?/i)?.[1],
    text.match(/\b(?:reading|read)\s+([A-Z][A-Za-z0-9',\- ]{2,70})(?:\s+(?:Foundational Skills|Vocabulary|Reading Workshop|Writing Workshop|High Frequency|Spelling)|[.])/i)?.[1],
    /Sky Color/i.test(text) ? "Sky Color" : "",
    /Young Frank,?\s*Architect/i.test(text) ? "Young Frank, Architect" : ""
  ].filter(Boolean);
  const title = candidates.find((item) => item && !/workshop|comprehension|vocabulary|spelling|grammar|skill/i.test(item));
  return title ? title.trim().replace(/\s+/g, " ").replace(/[:\-]+$/, "") : "";
}

function extractReadingSelectionsFromWeeklyText(text) {
  const source = String(text || "");
  const selections = [];
  const readAloudBlock = source.match(/\bREAD ALOUD\b\s*([\s\S]*?)(?:\bMY BOOK\b|\bFocus\b|\bReading Workshop\b|\bFoundational\b|\bHigh Frequency\b|\bSpelling\b|$)/i)?.[1] || "";
  const myBookBlock = source.match(/\bMY BOOK\b\s*([\s\S]*?)(?:\bFocus\b|\bReading Workshop\b|\bFoundational\b|\bHigh Frequency\b|\bSpelling\b|$)/i)?.[1] || "";
  [readAloudBlock, myBookBlock].forEach((block) => {
    String(block || "")
      .split(/\n|(?:^|\s)\d+\.\s*/)
      .map((item) => item.trim().replace(/^[^A-Za-z]+|[^A-Za-z?!',\- ]+$/g, ""))
      .filter((item) => /^[A-Z][A-Za-z?!',\- ]{2,80}$/.test(item))
      .filter((item) => !/^(?:read aloud|my book|focus|parents)$/i.test(item))
      .forEach((item) => selections.push(item.replace(/\s+/g, " ")));
  });
  const storyTitle = parseReadingTitleFromWeeklyText(source);
  return uniqueWords([storyTitle, ...selections].filter(Boolean)).slice(0, 8);
}

function extractHomeLearningSuggestions(text) {
  const block = String(text || "").match(/\bHome Learning Suggestions\b\s*([\s\S]*?)(?:\bOCR_PARSER_VERSION\b|\bSOURCE_PAGE_TYPE\b|\bRIGHT_COLUMN_WORDS\b|$)/i)?.[1] || "";
  return uniqueWords(
    block
      .split(/\n|•|·/)
      .map((item) => item.trim().replace(/\s+/g, " "))
      .filter((item) => item.length >= 12 && item.length <= 180)
      .filter((item) => /[A-Za-z]/.test(item))
  ).slice(0, 4);
}

function extractWeeklyLearningOutcomes(text, sourcePageType) {
  const source = String(text || "");
  if (!["uoiSubject", "englishSubject"].includes(sourcePageType)) return [];
  const block = source.match(/\bStudents will\b\s*([\s\S]*?)(?:\bHome Learning Suggestions\b|\bBilingual Vocabulary\b|\bOCR_PARSER_VERSION\b|\bSOURCE_PAGE_TYPE\b|$)/i)?.[1] || "";
  return uniqueWords(
    block
      .split(/\n|•|·/)
      .map((item) => item.trim().replace(/\s+/g, " "))
      .filter((item) => item.length >= 12 && item.length <= 220)
      .filter((item) => /^[A-Za-z]/.test(item))
      .filter((item) => !/\b(?:[A-Z]{2,}\b.*){3,}/.test(item))
      .filter((item) => !/\b(?:filifi|fuk|BRET|FHF|CHEF|HBL|REA AT EAST)\b/i.test(item))
      .filter((item) => !/\d+\.\d+/.test(item))
      .filter((item) => !/^Bilingual Vocabulary$/i.test(item))
  ).slice(0, 8);
}

function weeklyPlanPageType(flatText) {
  const text = String(flatText || "");
  const markerType = structuredOcrMarker(text, "SOURCE_PAGE_TYPE").split(/\s+/)[0] || "";
  if (["uoiSubject", "englishSubject", "sequenceSupplement", "unknown"].includes(markerType)) return markerType;
  if (/\bEnglish\b|英语|Home Learning Suggestions|read and discuss the story/i.test(text)) return "englishSubject";
  if (/\bUOI\b|探究|Bilingual Integration PYP Weekly Planner/i.test(text)) return "uoiSubject";
  if (/Weekly Sequence|Scope and Sequence|HMH Into Reading/i.test(text)) return "sequenceSupplement";
  return "unknown";
}

function parseWeeklyPlanData(text) {
  const rawNormalized = normalizeOcrText(text);
  const normalized = weeklyPlanEnglishText(text);
  const flat = flattenOcrText(normalized);
  const rawFlat = flattenOcrText(rawNormalized);
  const sourcePageType = weeklyPlanPageType(flat);
  const rightColumnText = structuredOcrBlock(rawNormalized, "RIGHT_COLUMN_WORDS", "END_RIGHT_COLUMN_WORDS");
  const trustedRightColumnEntries = trustedBilingualEntries(rightColumnText, flat, sourcePageType, normalized);
  const uoiFallbackEntries = sourcePageType === "uoiSubject" ? uoiSubjectVocabularyFallback(normalized) : [];
  const englishFallbackEntries = sourcePageType === "englishSubject" ? englishSubjectVocabularyFallback(normalized) : [];
  const uoiVocabularyEntries = sourcePageType === "uoiSubject"
    ? repairUoiBilingualOrder(trustedRightColumnEntries, uoiFallbackEntries).slice(0, 10)
    : [];
  const englishVocabularyEntries = sourcePageType === "englishSubject"
    ? (englishFallbackEntries.length >= 10 ? englishFallbackEntries : trustedRightColumnEntries).slice(0, 18)
    : [];
  const vocabularyText = extractWeeklySection(normalized, ["vocabulary"], ["generative", "uoi", "englishVocabulary", "englishPower", "highFrequency", "spelling", "comprehension", "foundational", "grammar", "jumpRope", "homeLearning", "dayActivity", "included"]);
  const uoiText = extractWeeklySection(normalized, ["uoi"], ["homeLearning", "comprehension", "foundational", "highFrequency", "spelling", "grammar", "jumpRope", "englishVocabulary", "englishPower", "dayActivity", "included"]);
  const englishVocabularyText = extractWeeklySection(normalized, ["englishVocabulary"], ["generative", "uoi", "englishPower", "highFrequency", "spelling", "comprehension", "foundational", "grammar", "jumpRope", "homeLearning", "dayActivity", "included"]);
  const englishPowerText = extractWeeklySection(normalized, ["englishPower"], ["generative", "uoi", "englishVocabulary", "highFrequency", "spelling", "comprehension", "foundational", "grammar", "jumpRope", "homeLearning", "dayActivity", "included"]);
  const highFrequencyText = extractWeeklySection(normalized, ["highFrequency"], ["foundational", "spelling", "basicWords", "review", "challenge", "jumpRope", "grammar", "dayActivity", "included"]);
  const spellingText = extractWeeklySection(normalized, ["spelling", "basicWords"], ["highFrequency", "foundational", "jumpRope", "grammar", "dayActivity", "included"]);
  const foundationalText = extractWeeklySection(normalized, ["foundational"], ["highFrequency", "spelling", "basicWords", "review", "challenge", "jumpRope", "grammar", "dayActivity", "included"]);
  const comprehensionText = extractWeeklySection(normalized, ["comprehension"], ["foundational", "highFrequency", "spelling", "grammar", "jumpRope", "dayActivity", "included"]);
  const grammarText = extractWeeklySection(normalized, ["grammar"], ["dayActivity", "included"]);
  const readerMatch = flat.match(/Jump Rope Reader:?\s*([\s\S]*?)(?:\s+Writing Workshop|\s+\d+\.\s*Opinion|\s+SOURCE_PAGE_TYPE|$)/i);
  const englishSubjectSpellingText = sourcePageType === "englishSubject" ? extractEnglishSubjectSpellingSections(normalized) : { basic: "", review: "", challenge: "" };
  const basicText = englishSubjectSpellingText.basic ||
    extractWeeklySection(normalized, ["basicWords"], ["review", "challenge", "highFrequency", "foundational", "spelling", "jumpRope", "grammar", "dayActivity", "included"]) ||
    extractBetween(spellingText, /\bBASIC WORDS?\s*:?\s*/i, [/\bREVIEW\s*:/i, /\bCHALLENGE\s*:/i]) ||
    (sourcePageType === "englishSubject" ? extractBetween(normalized, /\bbasic words?\b\s*[:,]?\s*/i, [/\breview words?\b/i, /\bchallenging words?\b/i, /\breview syllable/i, /\bread the jump/i, /\bHome Learning/i]) : "");
  const reviewText = englishSubjectSpellingText.review ||
    extractWeeklySection(normalized, ["review"], ["challenge", "basicWords", "highFrequency", "foundational", "spelling", "jumpRope", "grammar", "dayActivity", "included"]) ||
    extractBetween(spellingText, /\bREVIEW\s*:?\s*/i, [/\bCHALLENGE\s*:/i, /\bBASIC WORDS?\s*:/i, /\bHigh Frequency Words\s*:/i, /\bFoundational Skills\s*:/i, /Jump Rope/i, /Writing Workshop/i]) ||
    (sourcePageType === "englishSubject" ? extractBetween(normalized, /\breview words?\b\s*[:,]?\s*/i, [/\bchallenging words?\b/i, /\breview syllable/i, /\bread the jump/i, /\bHome Learning/i]) : "");
  const challengeText = englishSubjectSpellingText.challenge ||
    extractWeeklySection(normalized, ["challenge"], ["basicWords", "review", "highFrequency", "foundational", "spelling", "jumpRope", "grammar", "dayActivity", "included"]) ||
    extractBetween(spellingText, /\bCHALLENGE\s*:?\s*/i, [/\bBASIC WORDS?\s*:/i, /\bREVIEW\s*:/i, /\bHigh Frequency Words\s*:/i, /\bFoundational Skills\s*:/i, /Jump Rope/i, /Writing Workshop/i]) ||
    (sourcePageType === "englishSubject" ? extractBetween(normalized, /\bchallenging words?\b\s*[:,]?\s*/i, [/\breview syllable/i, /\bread the jump/i, /\bHome Learning/i]) : "");
  const phonicsText = `${foundationalText} ${spellingText}`;
  const highFrequencyWords = sourcePageType === "sequenceSupplement"
    ? highFrequencyWordsFromSection(highFrequencyText, flat)
    : [];
  const basicWords = cleanSpellingGroupWords(spellingWordsFromSection(basicText, 16), { limit: 12 });
  const blockedChallengeWords = new Set([...basicWords, ...highFrequencyWords].map(normalizeWord));
  const challengeWords = cleanSpellingGroupWords(spellingWordsFromSection(challengeText, 16), {
    limit: 4,
    blocked: blockedChallengeWords,
    allowHighFrequency: false
  });
  const blockedReviewWords = new Set([...basicWords, ...challengeWords, ...highFrequencyWords].map(normalizeWord));
  const reviewWords = cleanSpellingGroupWords(spellingWordsFromSection(reviewText, 16), {
    limit: 4,
    blocked: blockedReviewWords
  });
  const grammarPracticeWords = uniqueWords([...basicWords, ...reviewWords].filter(isContractionWord));
  const prefixPracticeWords = uniqueWords([
    ...(/prefix|re-|redo/i.test(flat) ? ["re-", "redo"] : []),
    ...(/prefix|un-|untie|uncover|unhappy/i.test(flat) ? ["un-", "untie", "uncover", "unhappy"] : []),
    ...knownWordsFromText(flat, ["balloon", "shampoo", "boyhood"])
  ]);
  const parsed = {
    sourcePageType,
    readingTitle: sourcePageType === "sequenceSupplement" ? "" : parseReadingTitleFromWeeklyText(flat),
    readingSelections: sourcePageType === "sequenceSupplement" ? [] : extractReadingSelectionsFromWeeklyText(normalized),
    jumpRopeReader: readerMatch?.[1]?.trim().replace(/\s+/g, " ") ?? "",
    vocabulary: sourcePageType === "unknown" ? wordsFromVocabularyText(vocabularyText) : [],
    uiVocabulary: sourcePageType === "uoiSubject"
      ? (uoiVocabularyEntries.length ? uoiVocabularyEntries : phrasesFromSection(uoiText))
      : [],
    englishVocabulary: sourcePageType === "englishSubject"
      ? (englishVocabularyEntries.length ? englishVocabularyEntries : wordsFromEnglishVocabularyText(`${englishVocabularyText}\n${englishPowerText}`))
      : [],
    uoiVocabulary: sourcePageType === "uoiSubject"
      ? (uoiVocabularyEntries.length ? uoiVocabularyEntries : phrasesFromSection(uoiText))
      : [],
    phonicsVocabulary: knownWordsFromText(phonicsText, phonicsExampleWords),
    highFrequencyWords,
    spelling: {
      skill: /\/oo\//i.test(spellingText) ? "Vowel Patterns /oo/" : "",
      basic: basicWords.filter((word) => !isContractionWord(word)),
      review: reviewWords.filter((word) => !isContractionWord(word)),
      challenge: challengeWords
    },
    spellingBasicWords: basicWords.filter((word) => !isContractionWord(word)),
    spellingReviewWords: reviewWords.filter((word) => !isContractionWord(word)),
    spellingChallengeWords: challengeWords,
    fieldSources: {
      uiVocabulary: sourcePageType === "uoiSubject" ? "bilingual-column" : "",
      englishVocabulary: sourcePageType === "englishSubject" ? "bilingual-column" : "",
      spelling: sourcePageType === "englishSubject" ? "english-subject-spelling" : sourcePageType === "sequenceSupplement" ? "sequence-supplement" : ""
    },
    grammarPracticeWords,
    wordStudyPracticeWords: prefixPracticeWords,
    foundationalSkills: uniqueWords([
      ...(/prefix|re-|un-/i.test(foundationalText) || /prefix|re-|un-/i.test(flat) ? ["prefixes"] : []),
      ...(/syllable/i.test(foundationalText) ? ["syllables"] : []),
      ...(/contraction/i.test(foundationalText) ? ["contractions"] : []),
      ...(/suffix/i.test(foundationalText) ? ["suffixes"] : []),
      ...(/segment/i.test(`${foundationalText} ${flat}`) ? ["segment phonemes"] : []),
      ...(/blend/i.test(`${foundationalText} ${flat}`) ? ["blend phonemes"] : []),
      ...(/manipulate/i.test(`${foundationalText} ${flat}`) ? ["manipulate phonemes"] : []),
      ...(/phoneme/i.test(`${foundationalText} ${flat}`) ? ["phonemes"] : []),
      ...(/diphthongs?\s+ow/i.test(flat) ? ["diphthongs ow, ou"] : []),
      ...(/diphthongs?\s+oy/i.test(flat) ? ["diphthongs oy, oi"] : [])
    ]),
    readingComprehension: uniqueWords([
      ...(/central idea|main idea/i.test(comprehensionText) ? ["central idea"] : []),
      ...(/retell/i.test(comprehensionText) ? ["retell"] : []),
      ...(/setting/i.test(comprehensionText) ? ["setting"] : []),
      ...(/genre characteristics|narrative nonfiction/i.test(flat) ? ["narrative nonfiction"] : []),
      ...(/author'?s purpose/i.test(flat) ? ["author's purpose"] : []),
      ...(/text organization/i.test(flat) ? ["text organization"] : [])
    ]),
    grammar: uniqueWords([
      ...(/adverb/i.test(grammarText) ? ["adverbs"] : []),
      ...(/adjective/i.test(grammarText) ? ["adjectives that compare"] : []),
      ...(/informational text/i.test(grammarText) ? ["informational text"] : []),
      ...(/biographical essay/i.test(grammarText) ? ["biographical essay"] : []),
      ...(/opinion letter/i.test(grammarText) || /opinion letter/i.test(flat) ? ["opinion letter"] : []),
      ...(/grammar review/i.test(grammarText) || /grammar review/i.test(flat) ? ["grammar review"] : []),
      ...(/contraction/i.test(flat) ? ["contractions"] : [])
    ]),
    homeLearningSuggestions: sourcePageType === "englishSubject" ? extractHomeLearningSuggestions(normalized) : [],
    weeklyLearningOutcomes: extractWeeklyLearningOutcomes(normalized, sourcePageType),
    pageTextSignals: uniqueWords([
      ...(/picture day/i.test(rawFlat) ? ["Picture Day"] : []),
      ...(/jump rope reader/i.test(rawFlat) ? ["Jump Rope Reader"] : [])
    ])
  };
  return applyWeeklyPlanOcrCorrections(parsed, flat);
}

function weeklyPlanDataHasContent(data = {}) {
  return Boolean(
    data.readingTitle ||
    (data.readingSelections ?? []).length ||
    data.jumpRopeReader ||
    data.phonicsFocus ||
    data.grammarFocus ||
    (data.foundationalSkills ?? []).length ||
    (data.readingComprehension ?? []).length ||
    (data.grammar ?? []).length ||
    (data.highFrequencyWords ?? []).length ||
    (data.spellingBasicWords ?? []).length ||
    (data.spellingReviewWords ?? []).length ||
    (data.spellingChallengeWords ?? []).length ||
    (data.uiVocabulary ?? []).length ||
    (data.englishVocabulary ?? []).length ||
    (data.vocabulary ?? []).length ||
    (data.homeLearningSuggestions ?? []).length ||
    (data.weeklyLearningOutcomes ?? []).length
  );
}

function extractCandidateWords(text) {
  const rawWords = text.match(/[A-Za-z][A-Za-z']{2,}/g) ?? [];
  const counts = rawWords.reduce((acc, word) => {
    const key = normalizeWord(word);
    if (key.length < 3) return acc;
    acc[key] = (acc[key] ?? 0) + 1;
    return acc;
  }, {});
  const ranked = Object.entries(counts)
    .filter(([word, count]) => highValueWords.has(word) || count >= 2 || word.length >= 6)
    .sort((a, b) => Number(highValueWords.has(b[0])) - Number(highValueWords.has(a[0])) || b[1] - a[1] || a[0].localeCompare(b[0]))
    .map(([word]) => word);
  return uniqueWords(ranked).slice(0, 24);
}

function extractTopicTags(text, material) {
  const topicMap = [
    ["simple machines", /simple machine|lever|wheel|ramp|pulley|wedge|screw|push|pull|lift/i],
    ["architecture", /architect|architecture|building|model|floor|designed|design/i],
    ["science", /observe|cause|effect|experiment|machine|force|motion/i],
    ["reading comprehension", /main idea|detail|sequence|retell|passage/i],
    ["grammar", /pronoun|verb|punctuation|contraction|sentence/i],
    ["phonics", /vowel|syllable|decode|blend|sound/i]
  ];
  return uniqueWords([...material.topicTags, ...topicMap.filter(([, pattern]) => pattern.test(text)).map(([tag]) => tag)]);
}

function cleanMainStoryText(text) {
  return englishOnlyOcrText(text)
    .replace(/\b(?:ENGLISH_ONLY_TEXT|END_ENGLISH_ONLY_TEXT|OCR_PARSER_VERSION|SOURCE_PAGE_TYPE|RIGHT_COLUMN_WORDS|END_RIGHT_COLUMN_WORDS)\b/gi, " ")
    .replace(/\b(?:HMH|Into Reading|Grade|Module|Page|Copyright|Houghton Mifflin Harcourt)\b/gi, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function mainStorySentencesFromText(text) {
  return uniqueWords(
    cleanMainStoryText(text)
      .split(/(?<=[.!?])\s+|\n+/)
      .map((item) => item.trim().replace(/\s+/g, " "))
      .filter((item) => item.length >= 18 && item.length <= 220)
      .filter((item) => /[A-Za-z]/.test(item))
      .filter((item) => !/^(?:read|answer|question|vocabulary|spelling|grammar|directions)\b/i.test(item))
  ).slice(0, 18);
}

function mainStoryCharactersFromSentences(sentences = [], title = "") {
  const blocked = new Set([
    ...String(title || "").split(/\s+/).map(normalizeWord),
    "the", "this", "that", "then", "first", "next", "last", "after", "before",
    "hmh", "into", "reading", "grade", "module", "page"
  ]);
  const counts = new Map();
  sentences.forEach((sentence) => {
    (sentence.match(/\b[A-Z][a-z']{2,}\b/g) ?? []).forEach((word) => {
      const key = normalizeWord(word);
      if (blocked.has(key)) return;
      counts.set(displayWord(word), (counts.get(displayWord(word)) ?? 0) + 1);
    });
  });
  return [...counts.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .map(([word]) => word)
    .slice(0, 5);
}

function extractMainStoryDetailsFromText(text, title = "") {
  const sentences = mainStorySentencesFromText(text);
  const evidence = sentences.slice(0, 8);
  const problem = sentences.find((sentence) => /\b(?:problem|challenge|could not|couldn't|can't|cannot|need|needs|want|wants|try|tries|hard)\b/i.test(sentence)) || "";
  const setting = sentences.find((sentence) => /\b(?:at|in|on|near|inside|outside|school|class|home|park|room|town)\b/i.test(sentence)) || "";
  const ending = sentences.length ? sentences[sentences.length - 1] : "";
  const bigIdea = sentences.find((sentence) => /\b(?:learn|learned|future|help|change|because|important|idea|purpose)\b/i.test(sentence)) || evidence[0] || "";
  const hasEvidence = evidence.join(" ").length >= 40;
  return {
    readingArticleSource: "hmh-into-reading",
    readingArticleContentStatus: hasEvidence ? "uploaded-pages" : "missing",
    readingArticleEvidence: evidence.join("\n"),
    readingArticleCharacters: mainStoryCharactersFromSentences(sentences, title),
    readingArticleSetting: setting,
    readingArticleProblem: problem,
    readingArticleKeyEvents: evidence.slice(0, 5),
    readingArticleEnding: ending,
    readingArticleBigIdea: bigIdea
  };
}

function mergeMainStoryDetails(detailsList = [], title = "") {
  const usable = detailsList.filter((details) => details?.readingArticleEvidence);
  if (!usable.length) {
    return {
      readingArticleSource: "hmh-into-reading",
      readingArticleContentStatus: "missing",
      readingArticleEvidence: "",
      readingArticleCharacters: [],
      readingArticleSetting: "",
      readingArticleProblem: "",
      readingArticleKeyEvents: [],
      readingArticleEnding: "",
      readingArticleBigIdea: ""
    };
  }
  const evidenceLines = uniqueWords(usable.flatMap((details) => splitRetellEvidence(details.readingArticleEvidence)));
  const keyEvents = uniqueWords(usable.flatMap((details) => details.readingArticleKeyEvents ?? []));
  return {
    readingArticleSource: "hmh-into-reading",
    readingArticleContentStatus: "uploaded-pages",
    readingArticleEvidence: evidenceLines.slice(0, 10).join("\n"),
    readingArticleCharacters: uniqueWords(usable.flatMap((details) => details.readingArticleCharacters ?? [])).slice(0, 8),
    readingArticleSetting: usable.find((details) => details.readingArticleSetting)?.readingArticleSetting || "",
    readingArticleProblem: usable.find((details) => details.readingArticleProblem)?.readingArticleProblem || "",
    readingArticleKeyEvents: (keyEvents.length ? keyEvents : evidenceLines).slice(0, 6),
    readingArticleEnding: [...usable].reverse().find((details) => details.readingArticleEnding)?.readingArticleEnding || "",
    readingArticleBigIdea: usable.find((details) => details.readingArticleBigIdea)?.readingArticleBigIdea || evidenceLines[0] || "",
    readingArticleTitle: title || ""
  };
}

function hmhArticleSearchUrl(title = "") {
  const query = `${title || "HMH Into Reading story"} HMH Into Reading Grade 1`;
  return `https://www.google.com/search?q=${encodeURIComponent(query)}`;
}

function buildWeeklySuggestions(material) {
  const skills = material.detectedSkills ?? [];
  const weekly = cleanedWeeklyPlanData(material.weeklyPlanData ?? {});
  const storyDetails = material.storyDetails || (material.materialKind === "main_story_pages"
    ? extractMainStoryDetailsFromText(material.extractedText || "", weekly.readingTitle || material.title)
    : null);
  const spellingWords = uniqueWords([
    ...(weekly.spellingBasicWords ?? []),
    ...(weekly.spellingReviewWords ?? []),
    ...(weekly.spellingChallengeWords ?? [])
  ]);
  const coreVocabulary = uniqueWords([
    ...(weekly.uiVocabulary ?? []),
    ...(weekly.englishVocabulary ?? []),
    ...(weekly.uiVocabulary?.length || weekly.englishVocabulary?.length ? [] : weekly.vocabulary ?? []),
    ...(weekly.uiVocabulary?.length ? [] : weekly.uoiVocabulary ?? [])
  ].filter((word) => !foundationalLearningTerms.has(normalizeWord(word))));
  const wordStudyExamples = uniqueWords([
    ...(weekly.grammarPracticeWords ?? []),
    ...(weekly.wordStudyPracticeWords ?? [])
  ]);
  return {
    sourceMaterialId: material.id,
    sourceTitle: material.title,
    readingTitle: weekly.readingTitle || "",
    readingSelections: uniqueWords(weekly.readingSelections ?? []),
    jumpRopeReader: weekly.jumpRopeReader || "",
    masteryWords: uniqueWords([...spellingWords, ...coreVocabulary]),
    sightWords: uniqueWords(weekly.highFrequencyWords ?? []).slice(0, 18),
    highFrequencyWords: uniqueWords(weekly.highFrequencyWords ?? []),
    spelling: normalizeSpellingGroups(weekly.spelling ?? {
      basic: weekly.spellingBasicWords ?? [],
      review: weekly.spellingReviewWords ?? [],
      challenge: weekly.spellingChallengeWords ?? []
    }),
    uiVocabulary: wordsToVocabularyEntries(weekly.uiVocabulary ?? [], "uiVocabulary"),
    englishVocabulary: wordsToVocabularyEntries(weekly.englishVocabulary ?? [], "englishVocabulary"),
    vocabularyWords: coreVocabulary,
    phonicsFocus: uniqueWords([...(weekly.foundationalSkills ?? []), ...skills.filter((skill) => /vowel|syllable|prefix|suffix|contraction|phonics|decode|sight/.test(skill))]).join(", "),
    grammarFocus: uniqueWords([...(weekly.grammar ?? []), ...skills.filter((skill) => /pronoun|verb|punctuation|sentence|grammar|adjective|adverb|essay|informational/.test(skill))]).join(", "),
    grammarPracticeWords: wordStudyExamples,
    readingFocus: uniqueWords([...(weekly.readingComprehension ?? []), ...skills.filter((skill) => /main idea|central idea|detail|sequence|cause|retell|evidence|reading|setting/.test(skill))]).join(", "),
    homeLearningSuggestions: uniqueWords(weekly.homeLearningSuggestions ?? []),
    weeklyLearningOutcomes: uniqueWords(weekly.weeklyLearningOutcomes ?? []),
    pageTextSignals: uniqueWords(weekly.pageTextSignals ?? []),
    epicIdeas: uniqueWords([...(material.topicTags ?? []), ...(material.skillTags ?? [])]).slice(0, 8),
    readingArticleSource: storyDetails?.readingArticleSource || "hmh-into-reading",
    readingArticleSourceUrl: storyDetails?.readingArticleSourceUrl || hmhArticleSearchUrl(weekly.readingTitle || material.title),
    readingArticleContentStatus: storyDetails?.readingArticleContentStatus || "missing",
    readingArticleEvidence: storyDetails?.readingArticleEvidence || "",
    readingArticleCharacters: storyDetails?.readingArticleCharacters ?? [],
    readingArticleSetting: storyDetails?.readingArticleSetting || "",
    readingArticleProblem: storyDetails?.readingArticleProblem || "",
    readingArticleKeyEvents: storyDetails?.readingArticleKeyEvents ?? [],
    readingArticleEnding: storyDetails?.readingArticleEnding || "",
    readingArticleBigIdea: storyDetails?.readingArticleBigIdea || ""
  };
}

function mergeWeeklyPlanDataList(materials = []) {
  const appendWords = (current = [], next = []) => uniqueWords([...(current ?? []), ...(next ?? [])]);
  const parsedMaterials = materials.map((material) => {
    const reparsedData = material.extractedText ? parseWeeklyPlanData(material.extractedText) : null;
    return cleanedWeeklyPlanData(weeklyPlanDataHasContent(reparsedData) ? reparsedData : material.weeklyPlanData ?? {});
  });
  const orderedData = [
    ...parsedMaterials.filter((data) => data.sourcePageType === "uoiSubject"),
    ...parsedMaterials.filter((data) => data.sourcePageType === "englishSubject"),
    ...parsedMaterials.filter((data) => !["uoiSubject", "englishSubject"].includes(data.sourcePageType))
  ];
  const merged = {
    readingTitle: "",
    readingSelections: [],
    jumpRopeReader: "",
    vocabulary: [],
    uiVocabulary: [],
    englishVocabulary: [],
    uoiVocabulary: [],
    phonicsVocabulary: [],
    highFrequencyWords: [],
    spelling: { skill: "", basic: [], review: [], challenge: [] },
    spellingBasicWords: [],
    spellingReviewWords: [],
    spellingChallengeWords: [],
    grammarPracticeWords: [],
    wordStudyPracticeWords: [],
    foundationalSkills: [],
    readingComprehension: [],
    grammar: [],
    homeLearningSuggestions: [],
    weeklyLearningOutcomes: [],
    pageTextSignals: [],
    sourcePageTypes: uniqueWords(parsedMaterials.map((data) => data.sourcePageType || "unknown"))
  };
  const englishSubjectData = parsedMaterials.filter((data) => data.sourcePageType === "englishSubject");
  const englishSubjectSpelling = normalizeSpellingGroups({
    basic: uniqueWords(englishSubjectData
      .filter((data) => data.fieldSources?.spelling === "english-subject-spelling")
      .flatMap((data) => [...(data.spelling?.basic ?? []), ...(data.spellingBasicWords ?? [])])),
    review: uniqueWords(englishSubjectData
      .filter((data) => data.fieldSources?.spelling === "english-subject-spelling")
      .flatMap((data) => [...(data.spelling?.review ?? []), ...(data.spellingReviewWords ?? [])])),
    challenge: uniqueWords(englishSubjectData
      .filter((data) => data.fieldSources?.spelling === "english-subject-spelling")
      .flatMap((data) => [...(data.spelling?.challenge ?? []), ...(data.spellingChallengeWords ?? [])]))
  });

  orderedData.forEach((data) => {
    if (!merged.readingTitle && data.readingTitle) merged.readingTitle = data.readingTitle;
    merged.readingSelections = appendWords(merged.readingSelections, data.readingSelections);
    if (!merged.jumpRopeReader && data.jumpRopeReader) merged.jumpRopeReader = data.jumpRopeReader;
    // Priority 1: bilingual vocabulary columns are source-of-truth fields.
    if (data.sourcePageType === "englishSubject" && data.fieldSources?.englishVocabulary === "bilingual-column") {
      merged.englishVocabulary = appendWords(merged.englishVocabulary, data.englishVocabulary);
    }
    if (data.sourcePageType === "uoiSubject" && data.fieldSources?.uiVocabulary === "bilingual-column") {
      merged.uiVocabulary = appendWords(merged.uiVocabulary, data.uiVocabulary);
      merged.uoiVocabulary = appendWords(merged.uoiVocabulary, data.uoiVocabulary);
    }

    // Priority 2: English subject spelling groups are core spelling fields.
    if (data.fieldSources?.spelling === "english-subject-spelling" || !englishSubjectData.length) {
      merged.spellingBasicWords = appendWords(merged.spellingBasicWords, data.spellingBasicWords);
      merged.spellingReviewWords = appendWords(merged.spellingReviewWords, data.spellingReviewWords);
      merged.spellingChallengeWords = appendWords(merged.spellingChallengeWords, data.spellingChallengeWords);
    }

    // Priority 3: remaining pages can supplement, but never replace, core fields.
    merged.vocabulary = appendWords(merged.vocabulary, data.vocabulary);
    merged.phonicsVocabulary = appendWords(merged.phonicsVocabulary, data.phonicsVocabulary);
    merged.highFrequencyWords = appendWords(merged.highFrequencyWords, data.highFrequencyWords);
    merged.grammarPracticeWords = appendWords(merged.grammarPracticeWords, data.grammarPracticeWords);
    merged.wordStudyPracticeWords = appendWords(merged.wordStudyPracticeWords, data.wordStudyPracticeWords);
    merged.foundationalSkills = appendWords(merged.foundationalSkills, data.foundationalSkills);
    merged.readingComprehension = appendWords(merged.readingComprehension, data.readingComprehension);
    merged.grammar = appendWords(merged.grammar, data.grammar);
    merged.homeLearningSuggestions = appendWords(merged.homeLearningSuggestions, data.homeLearningSuggestions);
    merged.weeklyLearningOutcomes = appendWords(merged.weeklyLearningOutcomes, data.weeklyLearningOutcomes);
    merged.pageTextSignals = appendWords(merged.pageTextSignals, data.pageTextSignals);
    merged.spelling = normalizeSpellingGroups({
      skill: merged.spelling.skill || data.spelling?.skill || "",
      basic: englishSubjectSpelling.basic.length ? englishSubjectSpelling.basic : appendWords(merged.spelling.basic, [...(data.spelling?.basic ?? []), ...(data.spellingBasicWords ?? [])]),
      review: englishSubjectSpelling.review.length ? englishSubjectSpelling.review : appendWords(merged.spelling.review, [...(data.spelling?.review ?? []), ...(data.spellingReviewWords ?? [])]),
      challenge: englishSubjectSpelling.challenge.length ? englishSubjectSpelling.challenge : appendWords(merged.spelling.challenge, [...(data.spelling?.challenge ?? []), ...(data.spellingChallengeWords ?? [])])
    });
  });
  if (merged.highFrequencyWords.length < 5 && merged.englishVocabulary.length) {
    const decodableFromEnglishVocabulary = merged.englishVocabulary.length >= 12
      ? merged.englishVocabulary.slice(-6)
      : merged.englishVocabulary.filter((word) => likelyHighFrequencyWords.has(normalizeWord(word)));
    merged.highFrequencyWords = decodableFromEnglishVocabulary.slice(0, 10);
  }
  return cleanedWeeklyPlanData(merged);
}

function buildWeeklySuggestionsFromMaterials(materials = [], batch = state.materialBatch) {
  const weeklyPlanMaterials = materials.filter((material) => materialReadyForWeeklyPlan(material));
  const mainStoryMaterials = materials.filter((material) => material.materialKind === "main_story_pages" || material.storyDetails?.readingArticleEvidence);
  const sourceMaterials = weeklyPlanMaterials.length ? weeklyPlanMaterials : materials.filter((material) => material.materialKind !== "main_story_pages");
  const mergedWeeklyPlanData = mergeWeeklyPlanDataList(sourceMaterials);
  const sourceTitle = batch?.title || sourceMaterials.map((item) => item.title || item.fileName).filter(Boolean).join(" + ") || "Uploaded weekly plan";
  const titleForStory = mergedWeeklyPlanData.readingTitle || activePlan().readingArticleTitle || activePlan().readingTitle || sourceTitle;
  const storyDetails = mergeMainStoryDetails(
    mainStoryMaterials.map((material) => material.storyDetails || extractMainStoryDetailsFromText(material.extractedText || "", titleForStory)),
    titleForStory
  );
  const syntheticMaterial = {
    id: batch?.id || createId(),
    title: sourceTitle,
    detectedSkills: uniqueWords([...sourceMaterials, ...mainStoryMaterials].flatMap((item) => item.detectedSkills ?? [])),
    detectedWords: [],
    topicTags: uniqueWords([...sourceMaterials, ...mainStoryMaterials].flatMap((item) => item.topicTags ?? [])),
    skillTags: uniqueWords([...sourceMaterials, ...mainStoryMaterials].flatMap((item) => item.skillTags ?? [])),
    extractedText: "",
    weeklyPlanData: mergedWeeklyPlanData,
    storyDetails
  };
  return {
    ...buildWeeklySuggestions(syntheticMaterial),
    sourceTitle,
    sourceMaterialIds: uniqueWords([...sourceMaterials, ...mainStoryMaterials].map((item) => item.id)),
    applyToNextWeek: false
  };
}

function materialNeedsStructuredWeeklyReparse(material = {}) {
  return material.materialKind === "weekly_plan" &&
    material.mimeType?.startsWith("image/") &&
    material.fileBlob &&
    !new RegExp(`OCR_PARSER_VERSION:\\s*${WEEKLY_PLAN_OCR_VERSION}`, "i").test(material.extractedText || "");
}

async function refreshStructuredWeeklyPlanMaterial(material) {
  if (!materialNeedsStructuredWeeklyReparse(material)) return material;
  const extractedText = await extractMaterialText(material);
  const analysis = analyzeMaterialText(extractedText, material);
  const refreshed = {
    ...material,
    extractedText,
    detectedSkills: analysis.detectedSkills,
    detectedWords: analysis.detectedWords,
    questionTypes: analysis.questionTypes,
    weeklyPlanData: analysis.weeklyPlanData,
    topicTags: uniqueWords([...(material.topicTags ?? []), ...analysis.topicTags]),
    parseStatus: weeklyPlanDataHasContent(analysis.weeklyPlanData) ? "parsed_weekly_plan" : material.parseStatus,
    parseMessage: weeklyPlanDataHasContent(analysis.weeklyPlanData)
      ? "Weekly plan parsed. Review suggestions before applying."
      : material.parseMessage,
    updatedAt: new Date().toISOString()
  };
  await saveMaterial(refreshed);
  return refreshed;
}

async function createWeeklyPlanFromBatch() {
  if (!state.materialBatch?.materialIds?.length) {
    state.materialsStatus = "Upload weekly plan files first.";
    renderMaterials();
    return;
  }
  const materials = (await Promise.all(state.materialBatch.materialIds.map(getMaterial))).filter(Boolean);
  if (!materials.length) {
    state.materialsStatus = "Could not find the uploaded files in local storage.";
    renderMaterials();
    return;
  }
  const needsRefresh = materials.some(materialNeedsStructuredWeeklyReparse);
  if (needsRefresh) {
    state.materialsStatus = "Updating OCR structure for this weekly plan...";
    renderMaterials();
  }
  const refreshedMaterials = await Promise.all(materials.map(refreshStructuredWeeklyPlanMaterial));
  if (needsRefresh) {
    state.materials = state.materials.map((material) => refreshedMaterials.find((item) => item.id === material.id) || material);
  }
  state.weeklySuggestions = buildWeeklySuggestionsFromMaterials(refreshedMaterials, state.materialBatch);
  state.materialsStatus = `Review the generated weekly plan, then choose this week or next week.`;
  renderMaterials();
}

function weeklyReviewFieldValue(field) {
  return document.querySelector(`[data-weekly-review-field="${field}"]`)?.value ?? "";
}

function syncWeeklySuggestionEditsFromReview() {
  const suggestion = state.weeklySuggestions;
  if (!suggestion) return null;
  const highFrequencyWords = splitEditableList(weeklyReviewFieldValue("highFrequencyWords"));
  const skills = splitEditableList(weeklyReviewFieldValue("skills"));
  const readingSelections = splitEditableList(weeklyReviewFieldValue("readingSelections"));
  const poemSelections = normalizePoemSelections(parsePoemSelectionsFromReview(weeklyReviewFieldValue("poemSelections")));
  const readingArticleEvidence = splitRetellEvidence(weeklyReviewFieldValue("readingArticleEvidence")).join("\n");
  const readingArticleSourceUrl = weeklyReviewFieldValue("readingArticleSourceUrl").trim();
  const readingArticleCharacters = splitEditableList(weeklyReviewFieldValue("readingArticleCharacters"));
  const readingArticleSetting = weeklyReviewFieldValue("readingArticleSetting").trim().replace(/\s+/g, " ");
  const readingArticleProblem = weeklyReviewFieldValue("readingArticleProblem").trim().replace(/\s+/g, " ");
  const readingArticleKeyEvents = splitEditableLines(weeklyReviewFieldValue("readingArticleKeyEvents"));
  const readingArticleEnding = weeklyReviewFieldValue("readingArticleEnding").trim().replace(/\s+/g, " ");
  const readingArticleBigIdea = weeklyReviewFieldValue("readingArticleBigIdea").trim().replace(/\s+/g, " ");
  const uiVocabulary = wordsToVocabularyEntries(splitEditableList(weeklyReviewFieldValue("uiVocabulary")), "uiVocabulary");
  const englishVocabulary = wordsToVocabularyEntries(splitEditableList(weeklyReviewFieldValue("englishVocabulary")), "englishVocabulary");
  const spelling = normalizeSpellingGroups({
    ...(suggestion.spelling ?? {}),
    basic: splitEditableList(weeklyReviewFieldValue("spellingBasic")),
    review: splitEditableList(weeklyReviewFieldValue("spellingReview")),
    challenge: splitEditableList(weeklyReviewFieldValue("spellingChallenge"))
  });
  const vocabularyWords = uniqueWords([
    ...uiVocabulary.map((item) => item.word),
    ...englishVocabulary.map((item) => item.word)
  ]);
  const edited = {
    ...suggestion,
    uiVocabulary,
    englishVocabulary,
    spelling,
    highFrequencyWords,
    sightWords: highFrequencyWords.slice(0, 18),
    grammarPracticeWords: splitEditableList(weeklyReviewFieldValue("grammarPracticeWords")),
    readingSelections,
    readingTitle: readingSelections[0] || "",
    readingArticleGenre: poemSelections.length ? "poetry" : (suggestion.readingArticleGenre || ""),
    poemSelections,
    readingArticleSource: "hmh-into-reading",
    readingArticleSourceUrl,
    readingArticleContentStatus: readingArticleEvidence || poemSelections.some((poem) => poem.evidence) ? "parent-edited" : "missing",
    readingArticleEvidence,
    readingArticleCharacters,
    readingArticleSetting,
    readingArticleProblem,
    readingArticleKeyEvents,
    readingArticleEnding,
    readingArticleBigIdea,
    weeklyLearningOutcomes: splitEditableLines(weeklyReviewFieldValue("weeklyLearningOutcomes")),
    homeLearningSuggestions: [],
    phonicsFocus: skills.join(", "),
    grammarFocus: skills.join(", "),
    readingFocus: skills.join(", "),
    vocabularyWords,
    masteryWords: uniqueWords([
      ...(spelling.basic ?? []),
      ...(spelling.review ?? []),
      ...(spelling.challenge ?? []),
      ...vocabularyWords
    ])
  };
  state.weeklySuggestions = edited;
  return edited;
}

function mergeWeeklyVocabulary(existing = [], words = []) {
  const items = Array.isArray(existing) ? existing : [];
  const byWord = new Map(items.filter((item) => !foundationalLearningTerms.has(normalizeWord(item.word))).map((item) => [normalizeWord(item.word), item]));
  words.forEach((word) => {
    const key = normalizeWord(word);
    if (!key || foundationalLearningTerms.has(key) || byWord.has(key)) return;
    byWord.set(key, {
      word,
      definition: wordHelp[key]?.[0] ?? "a word from this week's reading",
      example: wordHelp[key]?.[1] ?? `Use ${word} in a complete sentence.`
    });
  });
  return [...byWord.values()];
}

function vocabularyItemsFromWords(words = []) {
  return uniqueWords(words)
    .filter((word) => {
      const key = normalizeWord(word);
      return key && !foundationalLearningTerms.has(key);
    })
    .map((word) => {
      const key = normalizeWord(word);
      return {
        word: displayWord(word),
        definition: wordHelp[key]?.[0] ?? "a word from this week's learning",
        example: wordHelp[key]?.[1] ?? `Use ${displayWord(word)} in a complete sentence.`
      };
    });
}

function testBankWords(types = []) {
  const allowed = new Set(types);
  return Object.values(store.wordTestBank ?? {})
    .filter((entry) => !allowed.size || allowed.has(entry.wordType))
    .sort((a, b) => (b.lastSeenAt || "").localeCompare(a.lastSeenAt || ""))
    .map((entry) => entry.word);
}

function weeklyArchiveWordGroups(plan = activePlan()) {
  const vocabulary = weeklyVocabulary().map((item) => item.word);
  const uoiWords = vocabulary.filter((word) => knownUoiVocabulary.map(normalizeWord).includes(normalizeWord(word)));
  const phonicsWords = vocabulary.filter((word) => phonicsExampleWords.map(normalizeWord).includes(normalizeWord(word)));
  return {
    highFrequencyWords: uniqueWords([...(plan.sightWords ?? [])]),
    spellingWords: uniqueWords([...(plan.masteryWords ?? [])]),
    grammarPracticeWords: uniqueWords([...(plan.grammarPracticeWords ?? [])]),
    uoiVocabulary: uniqueWords(uoiWords),
    topicVocabulary: uniqueWords(vocabulary.filter((word) => !uoiWords.map(normalizeWord).includes(normalizeWord(word)) && !phonicsWords.map(normalizeWord).includes(normalizeWord(word)))),
    phonicsReviewWords: uniqueWords(phonicsWords),
    weeklyVocabulary: uniqueWords(vocabulary)
  };
}

function learnerArchiveSnapshot(learner) {
  const progress = store.learners[learner.id] ?? getDefaultProgress();
  return {
    learnerId: learner.id,
    learnerName: profileName(learner.id),
    completedToday: progress.centerCompletions?.[todayKey()]?.length ?? 0,
    weakSkills: [...(progress.weakSkills ?? [])],
    spellingMisses: Object.values(progress.wrongWordProfile ?? {})
      .sort((a, b) => (b.missCount ?? 0) - (a.missCount ?? 0))
      .slice(0, 20),
    sightWordMisses: Object.values(progress.sightWordProfile ?? {})
      .filter((item) => (item.missed ?? 0) > 0)
      .sort((a, b) => (b.missCount ?? b.missed ?? 0) - (a.missCount ?? a.missed ?? 0))
      .slice(0, 20)
  };
}

function addWordsToTestBank(archive, groups) {
  const typedGroups = [
    ["sight", groups.highFrequencyWords],
    ["spelling", groups.spellingWords],
    ["grammar", groups.grammarPracticeWords],
    ["uoi", groups.uoiVocabulary],
    ["topicVocabulary", groups.topicVocabulary],
    ["phonicsReview", groups.phonicsReviewWords],
    ["mastery", groups.weeklyVocabulary]
  ];
  if (!store.wordTestBank) store.wordTestBank = {};
  typedGroups.forEach(([wordType, words]) => {
    (words ?? []).forEach((word) => {
      const key = normalizeWord(word);
      if (!key || foundationalLearningTerms.has(key)) return;
      const current = store.wordTestBank[key] ?? {
        word: displayWord(word),
        wordType,
        firstSeenAt: new Date().toISOString(),
        sourceArchives: []
      };
      const sourceArchives = uniqueWords([...(current.sourceArchives ?? []), archive.archiveLabel]);
      store.wordTestBank[key] = {
        ...current,
        word: current.word || displayWord(word),
        wordType: current.wordType === "mastery" ? wordType : current.wordType,
        definition: wordHelp[key]?.[0] ?? current.definition ?? "a word from weekly learning",
        example: wordHelp[key]?.[1] ?? current.example ?? `Use ${displayWord(word)} in a complete sentence.`,
        sourceArchiveLabel: archive.archiveLabel,
        sourceCalendarWeek: archive.calendarWeek,
        sourceCalendarYear: archive.calendarYear,
        sourceArchives,
        lastSeenAt: new Date().toISOString()
      };
    });
  });
}

function archiveCurrentWeek(parentNotes = "", { render = true } = {}) {
  const plan = activePlan();
  const week = planWeekInfo(plan);
  const groups = weeklyArchiveWordGroups(plan);
  const archive = {
    id: `${week.calendarYear}-cw-${String(week.calendarWeek).padStart(2, "0")}`,
    ...week,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    schoolWeekLabel: plan.schoolWeekLabel || "",
    sourceMaterials: state.materials.map((material) => ({ id: material.id, title: material.title })).slice(0, 12),
    weeklyPlanSnapshot: JSON.parse(JSON.stringify(plan)),
    identifiedSkills: {
      phonics: plan.phonicsFocus || "",
      grammar: plan.grammarFocus || "",
      reading: plan.readingTitle || plan.readingArticleTitle || "",
      uoi: plan.uoiTopic || ""
    },
    identifiedWords: groups,
    learnerSnapshots: learners.map(learnerArchiveSnapshot),
    parentNotes
  };
  if (!store.weeklyArchives) store.weeklyArchives = [];
  const existingIndex = store.weeklyArchives.findIndex((item) => item.archiveLabel === archive.archiveLabel);
  if (existingIndex >= 0) {
    archive.createdAt = store.weeklyArchives[existingIndex].createdAt || archive.createdAt;
    store.weeklyArchives[existingIndex] = archive;
  } else {
    store.weeklyArchives.unshift(archive);
  }
  addWordsToTestBank(archive, groups);
  saveStore();
  if (render) renderAll();
  return archive;
}

function ensureCurrentWeekArchiveSeed() {
  const label = planWeekInfo(activePlan()).archiveLabel;
  if ((store.weeklyArchives ?? []).some((archive) => archive.archiveLabel === label)) return;
  archiveCurrentWeek("", { render: false });
}

function clearAllTestPracticeRecords() {
  learners.forEach((learner) => {
    store.learners[learner.id] = clearPracticeProgress(store.learners[learner.id]);
  });
  saveStore();
  renderAll();
}

function buildPlanFromSuggestion(basePlan, suggestion, weekInfo, status = "active") {
  const readingSelections = uniqueWords([...(basePlan.readingSelections ?? []), ...(suggestion.readingSelections ?? [])]);
  const hasPoemSelectionEdits = Object.prototype.hasOwnProperty.call(suggestion, "poemSelections");
  const suggestedPoemSelections = normalizePoemSelections(suggestion.poemSelections ?? []);
  const poemSelections = hasPoemSelectionEdits
    ? suggestedPoemSelections
    : normalizePoemSelections(basePlan.poemSelections ?? []);
  const weeklyLearningOutcomes = uniqueWords([...(basePlan.weeklyLearningOutcomes ?? []), ...(suggestion.weeklyLearningOutcomes ?? [])]);
  const homeLearningSuggestions = uniqueWords([...(basePlan.homeLearningSuggestions ?? []), ...(suggestion.homeLearningSuggestions ?? [])]);
  const pageTextSignals = uniqueWords([...(basePlan.pageTextSignals ?? []), ...(suggestion.pageTextSignals ?? [])]);
  const hasSuggestedStoryEvidence = Boolean(String(suggestion.readingArticleEvidence || "").trim());
  const hasSuggestedPoemEvidence = poemSelections.some((poem) => poem.evidence);
  const readingSkills = uniqueWords([
    ...(basePlan.readingSkills ?? []),
    ...String(suggestion.readingFocus || "").split(",").map((item) => item.trim()).filter(Boolean)
  ]);
  const plan = {
    ...basePlan,
    readingArticleTitle: suggestion.readingTitle || readingSelections[0] || basePlan.readingArticleTitle || suggestion.sourceTitle,
    readingTitle: suggestion.readingTitle || readingSelections[0] || basePlan.readingTitle || suggestion.sourceTitle,
    readingArticleGenre: suggestion.readingArticleGenre || basePlan.readingArticleGenre || "",
    readingArticleSummary: weeklyLearningOutcomes[0] || `Practice words and skills from ${suggestion.readingTitle || suggestion.sourceTitle}.`,
    readingSelections,
    poemSelections,
    readingSkills,
    weeklyLearningOutcomes,
    homeLearningSuggestions,
    pageTextSignals,
    readingArticleSource: suggestion.readingArticleSource || basePlan.readingArticleSource || "hmh-into-reading",
    readingArticleSourceUrl: suggestion.readingArticleSourceUrl || basePlan.readingArticleSourceUrl || "",
    readingArticleContentStatus: hasSuggestedStoryEvidence || hasSuggestedPoemEvidence
      ? (suggestion.readingArticleContentStatus || "parent-edited")
      : (basePlan.readingArticleContentStatus || "missing"),
    readingArticleEvidence: hasSuggestedStoryEvidence ? suggestion.readingArticleEvidence : (basePlan.readingArticleEvidence || ""),
    readingArticleCharacters: hasSuggestedStoryEvidence ? uniqueWords(suggestion.readingArticleCharacters ?? []) : uniqueWords(basePlan.readingArticleCharacters ?? []),
    readingArticleSetting: hasSuggestedStoryEvidence ? (suggestion.readingArticleSetting || "") : (basePlan.readingArticleSetting || ""),
    readingArticleProblem: hasSuggestedStoryEvidence ? (suggestion.readingArticleProblem || "") : (basePlan.readingArticleProblem || ""),
    readingArticleKeyEvents: hasSuggestedStoryEvidence ? uniqueWords(suggestion.readingArticleKeyEvents ?? []) : uniqueWords(basePlan.readingArticleKeyEvents ?? []),
    readingArticleEnding: hasSuggestedStoryEvidence ? (suggestion.readingArticleEnding || "") : (basePlan.readingArticleEnding || ""),
    readingArticleBigIdea: hasSuggestedStoryEvidence ? (suggestion.readingArticleBigIdea || "") : (basePlan.readingArticleBigIdea || ""),
    jumpRopeReader: suggestion.jumpRopeReader || basePlan.jumpRopeReader || "",
    phonicsFocus: uniqueWords([basePlan.phonicsFocus, suggestion.phonicsFocus].filter(Boolean).join(", ").split(",")).join(", "),
    grammarFocus: uniqueWords([basePlan.grammarFocus, suggestion.grammarFocus, suggestion.readingFocus].filter(Boolean).join(", ").split(",")).join(", "),
    masteryWords: uniqueWords([...(basePlan.masteryWords ?? []), ...suggestion.masteryWords]),
    sightWords: uniqueWords([...(basePlan.sightWords ?? []), ...suggestion.sightWords]).slice(0, 24),
    highFrequencyWords: uniqueWords([...(basePlan.highFrequencyWords ?? []), ...(suggestion.highFrequencyWords ?? [])]),
    spelling: normalizeSpellingGroups({
      skill: suggestion.spelling?.skill || basePlan.spelling?.skill || "",
      basic: uniqueWords([...(basePlan.spelling?.basic ?? []), ...(suggestion.spelling?.basic ?? [])]),
      review: uniqueWords([...(basePlan.spelling?.review ?? []), ...(suggestion.spelling?.review ?? [])]),
      challenge: uniqueWords([...(basePlan.spelling?.challenge ?? []), ...(suggestion.spelling?.challenge ?? [])])
    }),
    uiVocabulary: normalizeVocabularyEntries([...(basePlan.uiVocabulary ?? []), ...(suggestion.uiVocabulary ?? [])], "uiVocabulary"),
    englishVocabulary: normalizeVocabularyEntries([...(basePlan.englishVocabulary ?? []), ...(suggestion.englishVocabulary ?? [])], "englishVocabulary"),
    grammarPracticeWords: uniqueWords([...(basePlan.grammarPracticeWords ?? []), ...(suggestion.grammarPracticeWords ?? [])]),
    weeklyVocabulary: mergeWeeklyVocabulary(basePlan.weeklyVocabulary, suggestion.vocabularyWords ?? []),
    epicIdeas: uniqueWords([...(basePlan.epicIdeas ?? []), ...suggestion.epicIdeas]).slice(0, 24),
    sourceTitle: suggestion.sourceTitle,
    updatedAt: new Date().toISOString()
  };
  const planWithMetadata = applyWeekMetadata(plan, weekInfo, status);
  return {
    ...planWithMetadata,
    listeningStories: listeningStoriesForPlan(planWithMetadata)
  };
}

function seedReviewQueueForPlan(plan, sourceTitle = "weekly words") {
  learners.forEach((learner) => {
    planWeeklyNewWords(plan).forEach((word) => {
      const key = normalizeWord(word);
      if (!store.learners[learner.id].reviewQueue[key]) {
        store.learners[learner.id].reviewQueue[key] = { stage: 0, nextReview: todayKey(), source: sourceTitle };
      }
    });
  });
}

function applyWeeklySuggestions(target = "auto") {
  const suggestion = state.weeklySuggestions;
  if (!suggestion) return;
  ensureWeeklyPlanIndex();
  const stageNextWeek = target === "next-week" || (target === "auto" && (Boolean(suggestion.applyToNextWeek) || isFridayOrSaturday()));
  const weekInfo = stageNextWeek ? nextWeekInfoForDate() : learningWeekInfo();
  const emptyNextWeekPlan = { ...defaultWeeklyPlan, weekId: weekInfo.weekId, readingTitle: "", readingArticleTitle: "", readingArticleSummary: "", masteryWords: [], sightWords: [], grammarPracticeWords: [], weeklyVocabulary: [], epicIdeas: [], listeningStories: [] };
  const basePlan = stageNextWeek
    ? planMatchesWeekId(store.pendingWeeklyPlan, weekInfo.weekId)
      ? store.pendingWeeklyPlan
      : store.weeklyPlansByWeekId?.[weekInfo.weekId] || emptyNextWeekPlan
    : activePlan();
  const nextPlan = buildPlanFromSuggestion(basePlan, suggestion, weekInfo, stageNextWeek ? "pending" : "active");
  if (stageNextWeek) {
    store.pendingWeeklyPlan = nextPlan;
    store.weeklyPlansByWeekId[weekInfo.weekId] = nextPlan;
  } else {
    store.weeklyPlan = nextPlan;
    store.currentWeekId = nextPlan.weekId;
    store.weeklyPlansByWeekId[weekInfo.weekId] = nextPlan;
    seedReviewQueueForPlan(nextPlan, suggestion.sourceTitle);
  }
  uniqueWords([...(suggestion.masteryWords ?? []), ...(suggestion.vocabularyWords ?? [])]).forEach((word) => {
    const key = normalizeWord(word);
    if (key && !wordHelp[key]) {
      wordHelp[key] = ["a word from this week's learning material", `Try using ${key} in a complete sentence.`];
    }
  });
  state.weeklySuggestions = null;
  state.materialBatch = null;
  saveStore();
  state.materialsStatus = stageNextWeek
    ? `${suggestion.sourceTitle} staged for ${weekInfo.calendarWeekLabel}. It will activate Sunday.`
    : "Suggestions applied to this week's plan.";
  renderAll();
}

async function useMaterialForWeek(id) {
  const material = await getMaterial(id);
  if (!material) return;
  state.weeklySuggestions = buildWeeklySuggestions(material);
  state.materialsStatus = "Review suggestions, then apply them to this week.";
  renderMaterials();
}

async function changeParentPassword(event) {
  event.preventDefault();
  const current = document.querySelector("#currentParentPasswordInput").value;
  const next = document.querySelector("#newParentPasswordInput").value;
  const status = document.querySelector("#passwordStatus");
  if (!next || next.length < 6) {
    status.textContent = "Use at least 6 characters.";
    return;
  }
  if ((await hashText(current)) !== store.parentPasswordHash) {
    status.textContent = "Current password is incorrect.";
    return;
  }
  store.parentPasswordHash = await hashText(next);
  saveStore();
  event.target.reset();
  status.textContent = "Password changed.";
}

function stopAudio() {
  if (!("speechSynthesis" in window)) return;
  state.audioSequenceTimers.forEach((timer) => window.clearTimeout(timer));
  state.audioSequenceTimers = [];
  window.speechSynthesis.cancel();
  state.activeAudioKey = null;
  state.currentUtterance = null;
  document.querySelectorAll("[data-audio-target], [data-spelling-audio], #speakWordButton").forEach((button) => {
    if (button.dataset.idleLabel) button.textContent = button.dataset.idleLabel;
    button.classList.remove("audio-active");
  });
}

function availableVoices() {
  if (!("speechSynthesis" in window)) return [];
  return window.speechSynthesis.getVoices?.() ?? [];
}

function preferredEnglishVoice() {
  const voices = availableVoices();
  const femaleNames = /female|girl|child|serena|kate|martha|tessa|susan|samantha|victoria|karen|moira|fiona|ava|allison|joanna|salli|siri/i;
  const maleNames = /male|daniel|alex|oliver|arthur|fred|tom|ralph|george|lee|liam|aaron|nathan|david|mark|paul|james|john/i;
  const englishVoices = voices.filter((voice) => /^en/i.test(voice.lang));
  const britishVoices = englishVoices.filter((voice) => /^en-GB/i.test(voice.lang));
  const isFemaleLike = (voice) => femaleNames.test(voice.name);
  const isKnownMale = (voice) => maleNames.test(voice.name);
  const chosen =
    britishVoices.find((voice) => isFemaleLike(voice) && !isKnownMale(voice)) ||
    englishVoices.find((voice) => isFemaleLike(voice) && !isKnownMale(voice)) ||
    britishVoices.find((voice) => !isKnownMale(voice)) ||
    englishVoices.find((voice) => !isKnownMale(voice)) ||
    englishVoices[0] ||
    null;
  if (chosen) {
    window.readingBridgeSelectedVoice = {
      name: chosen.name,
      lang: chosen.lang,
      available: voices.map((voice) => `${voice.name} (${voice.lang})`)
    };
  }
  return chosen;
}

if ("speechSynthesis" in window) {
  window.speechSynthesis.onvoiceschanged = () => {
    preferredEnglishVoice();
  };
}

function prepareAudioButton(button) {
  if (!button) return;
  button.dataset.idleLabel = button.dataset.idleLabel || button.textContent;
  button.textContent = "Stop";
  button.classList.add("audio-active");
}

function playAudio(text, key, button = null, options = {}) {
  if (!("speechSynthesis" in window)) return;
  if (state.activeAudioKey === key) {
    stopAudio();
    return;
  }

  stopAudio();
  const utterance = new SpeechSynthesisUtterance(text);
  const voice = preferredEnglishVoice();
  if (voice) utterance.voice = voice;
  utterance.lang = voice?.lang || options.lang || "en-GB";
  utterance.rate = options.rate ?? 0.84;
  utterance.pitch = options.pitch ?? 1.3;
  state.activeAudioKey = key;
  state.currentUtterance = utterance;
  prepareAudioButton(button);
  utterance.addEventListener("end", stopAudio);
  utterance.addEventListener("error", stopAudio);
  window.speechSynthesis.speak(utterance);
}

function playAudioSequence(parts, key, button = null, options = {}) {
  if (!("speechSynthesis" in window)) return;
  if (state.activeAudioKey === key) {
    stopAudio();
    return;
  }
  stopAudio();
  state.activeAudioKey = key;
  prepareAudioButton(button);
    const voice = preferredEnglishVoice();
    let index = 0;
  const speakNext = () => {
    if (state.activeAudioKey !== key) return;
    const part = parts[index];
    if (!part) {
      stopAudio();
      return;
    }
    const utterance = new SpeechSynthesisUtterance(part.text);
    if (voice) utterance.voice = voice;
    utterance.lang = voice?.lang || options.lang || "en-GB";
    utterance.rate = options.rate ?? 0.84;
    utterance.pitch = options.pitch ?? 1.3;
    state.currentUtterance = utterance;
    utterance.addEventListener("end", () => {
      if (state.activeAudioKey !== key) return;
      index += 1;
      const timer = window.setTimeout(speakNext, part.pauseAfter ?? 0);
      state.audioSequenceTimers.push(timer);
    });
    utterance.addEventListener("error", stopAudio);
    window.speechSynthesis.speak(utterance);
  };
  speakNext();
}

function wordExample(word, { context = "" } = {}) {
  const key = normalizeWord(word);
  if (context === "spelling") {
    return spellingExampleSentences[key]
      ?? weeklyVocabularyEntry(key)?.example
      ?? wordHelp[key]?.[1]
      ?? `Listen for ${displayWord(word)} in a clear sentence.`;
  }
  return weeklyVocabularyEntry(key)?.example
    ?? wordHelp[key]?.[1]
    ?? spellingExampleSentences[key]
    ?? `The word ${displayWord(word)} tells about something in this week's lesson.`;
}

function learningEntryForWord(word) {
  const key = normalizeWord(word);
  const weeklyEntry = weeklyVocabularyEntry(key);
  if (weeklyEntry) return weeklyEntry;
  return {
    word: displayWord(word),
    definition: wordHelp[key]?.[0] ?? "a word good readers know and use",
    example: wordHelp[key]?.[1] ?? `I can use ${displayWord(word)} in a sentence.`
  };
}

function speakWordHelp(word, button = null, keyPrefix = "word-help") {
  const key = normalizeWord(word);
  const popover = document.querySelector("#wordPopover");
  const popoverWord = normalizeWord(document.querySelector("#popoverWord")?.textContent ?? "");
  const entry = learningEntryForWord(word);
  const definition = !popover?.hidden && popoverWord === key
    ? document.querySelector("#popoverDefinition")?.textContent || entry.definition
    : entry.definition;
  const example = !popover?.hidden && popoverWord === key
    ? document.querySelector("#popoverExample")?.textContent || entry.example
    : entry.example;
  playAudioSequence(
    [
      { text: entry.word || displayWord(word), pauseAfter: 700 },
      { text: definition, pauseAfter: 850 },
      { text: `Example: ${example}` }
    ],
    `${keyPrefix}-${key}`,
    button
  );
}

function speakSpellingPrompt(word, key = `spelling-${normalizeWord(word)}`, button = null) {
  playAudioSequence(
    [
      { text: displayWord(word), pauseAfter: 750 },
      { text: wordExample(word, { context: "spelling" }) }
    ],
    key,
    button,
    { rate: 0.82, pitch: 1.16 }
  );
}

function audioButton(label, text, key, className = "secondary-action") {
  return `<button class="${className}" type="button" data-audio-target="${escapeHTML(text)}" data-audio-key="${escapeHTML(key)}">${label}</button>`;
}

function scheduleWordReview(word, result) {
  const progress = activeProgress();
  const key = normalizeWord(word);
  const current = progress.wordStats[key] ?? { correct: 0, attempts: 0 };
  const correct = result === "correct";
  const nextCorrect = current.correct + (correct ? 1 : 0);
  const attempts = current.attempts + 1;
  const mastered = nextCorrect >= 4 && correct;
  progress.wordStats[key] = { correct: nextCorrect, attempts, mastered, lastResult: result, lastSeen: todayKey() };

  const intervalIndex = Math.min(nextCorrect, REVIEW_INTERVALS.length - 1);
  progress.reviewQueue[key] = {
    stage: intervalIndex,
    nextReview: mastered ? addDays(new Date(), REVIEW_INTERVALS[REVIEW_INTERVALS.length - 1]) : addDays(new Date(), REVIEW_INTERVALS[intervalIndex]),
    source: "weekly words"
  };

  saveStore();
  renderAll();
}

function renderLearners() {
  document.querySelector("#learnerButtons").innerHTML = learners
    .map(
      (learner) => `
        <button class="learner-button ${learner.id === state.activeLearnerId ? "active" : ""}" type="button" data-learner-id="${learner.id}">
          <strong>${escapeHTML(profileName(learner.id))}</strong>
          <span>${learner.target}</span>
        </button>
      `
    )
    .join("");
}

function renderProfileGate() {
  document.querySelector("#profileGrid").innerHTML = profiles
    .map((profile) => {
      const isEditing = state.editingProfileId === profile.id && profile.role === "child";
      const name = profileName(profile.id);
      const avatar = profile.avatarImage
        ? `<img class="profile-avatar-art" src="${profile.avatarImage}" alt="" aria-hidden="true" />`
        : profile.role === "child"
          ? `<span class="profile-letter-art" aria-hidden="true">${escapeHTML(profile.avatar || profile.defaultName.slice(0, 1))}</span>`
          : `<span class="parent-gate-mark" aria-hidden="true">Parent</span>`;
      const progress = store.learners?.[profile.id] ?? getDefaultProgress();
      return `
        <article class="profile-card ${profile.role} profile-${profile.id}">
          ${profile.role === "child" ? `<span class="profile-heart-badge" aria-label="${heartCount(progress)} hearts">♥ ${heartCount(progress)}</span>` : ""}
          ${avatar}
          <div class="profile-name-area">
            ${
              isEditing
                ? `<label>
                    <span class="sr-only">Name</span>
                    <input type="text" value="${escapeHTML(name)}" data-profile-name="${profile.id}" aria-label="${profile.defaultName} name" autofocus />
                  </label>
                  <button class="secondary-action" type="button" data-finish-name="${profile.id}">Done</button>`
                : `<div class="profile-name-row">
                    <button class="profile-name-button" type="button" data-start-profile="${profile.id}">
                      ${escapeHTML(name)}
                    </button>
                    ${
                      profile.role === "child"
                        ? `<button class="edit-name-button" type="button" data-edit-name="${profile.id}" aria-label="Edit ${escapeHTML(name)} name">✎</button>`
                        : ""
                    }
                  </div>`
            }
          </div>
        </article>
      `;
    })
    .join("");
}

function renderSummary() {
  const levels = getTodayAdventureLevels();
  const completed = levels.filter((level) => level.completed);
  document.querySelector("#todayStars").textContent = completed.length;
  const heartCountTarget = document.querySelector("#heartCount");
  if (heartCountTarget) heartCountTarget.textContent = heartCount(activeProgress());
  document.querySelector("#masteredCount").textContent = masteredWords().length;
  document.querySelector("#reviewDueCount").textContent = dueReviewWords().length;
  document.querySelector("#missionStatus").textContent = completed.length >= levels.length ? "Complete" : `${completed.length}/${levels.length}`;
}

function topicVocabularyWords(topic = "") {
  const skip = new Set(["and", "the", "with", "about", "topic", "vocabulary"]);
  return uniqueWords(
    String(topic || "")
      .split(/[^A-Za-z']+/)
      .map((word) => word.trim())
      .filter((word) => word.length > 2 && !skip.has(normalizeWord(word)))
  );
}

function wordToLearnGroups(plan = activePlan()) {
  const vocabulary = vocabularyEntriesForPlan(plan);
  return [
    {
      id: "uiVocabulary",
      title: "UOI / Topic Vocabulary",
      words: vocabulary.uiVocabulary
    },
    {
      id: "englishVocabulary",
      title: "English Vocabulary",
      words: vocabulary.englishVocabulary
    }
  ];
}

function wordToLearnEntries(plan = activePlan()) {
  const byWord = new Map();
  wordToLearnGroups(plan).forEach((group) => {
    group.words.forEach((entry) => {
      const key = normalizeWord(entry.word);
      if (key && !byWord.has(key)) byWord.set(key, { ...entry, category: entry.category || group.id });
    });
  });
  return [...byWord.values()];
}

function weeklyVocabulary() {
  const plan = activePlan();
  const structuredWords = wordToLearnEntries(plan);
  if (structuredWords.length) return structuredWords;

  const orderedItems =
    plan.weekId === CURRENT_WEEK_ID
      ? [...defaultWeeklyPlan.weeklyVocabulary]
      : Array.isArray(plan.weeklyVocabulary)
        ? [...plan.weeklyVocabulary]
        : [];
  uniqueWords([
    ...(plan.masteryWords ?? []),
    ...(plan.sightWords ?? []),
    ...(plan.grammarPracticeWords ?? []),
    ...topicVocabularyWords(plan.uoiTopic)
  ]).forEach((word) => {
    const key = normalizeWord(word);
    if (!key || orderedItems.some((item) => normalizeWord(item.word) === key)) return;
    orderedItems.push({
      word: displayWord(word),
      definition: wordHelp[key]?.[0] || "a word from this week's learning",
      example: wordHelp[key]?.[1] || `Use ${displayWord(word)} in a complete sentence.`
    });
  });
  const byWord = new Map();
  orderedItems.forEach((item) => {
    const key = normalizeWord(item.word);
    if (!key || foundationalLearningTerms.has(key) || byWord.has(key)) return;
    byWord.set(key, {
      word: item.word,
      definition: item.definition || wordHelp[key]?.[0] || "a word from this week's learning",
      example: item.example || wordHelp[key]?.[1] || `Use ${item.word} in a complete sentence.`
    });
  });
  return [...byWord.values()];
}

function weeklyVocabularyEntry(word) {
  const key = normalizeWord(word);
  return weeklyVocabulary().find((item) => normalizeWord(item.word) === key);
}

function renderChildWeeklyPlan() {
  const plan = activePlan();
  const vocabulary = weeklyVocabulary();
  const vocabularyGroups = wordToLearnGroups(plan).filter((group) => group.words.length);
  const storyTitle = plan.readingArticleTitle || plan.readingTitle;
  const skillsSummary = uniqueWords([
    plan.phonicsFocus || "phonics",
    plan.grammarFocus || "grammar",
    plan.uoiTopic || "topic vocabulary"
  ]).slice(0, 2).join(" + ");
  document.querySelector("#childWeeklyPlan").innerHTML = `
    <details class="weekly-plan-drawer">
      <summary class="weekly-plan-summary">
        <span class="weekly-plan-summary-icon" aria-hidden="true">
          <span class="book-shape"></span>
        </span>
        <span class="weekly-plan-summary-copy">
          <span class="eyebrow">This Week</span>
          <strong>${escapeHTML(storyTitle)}</strong>
          <small>${escapeHTML(skillsSummary)} · ${vocabulary.length} words</small>
        </span>
        <span class="weekly-plan-summary-action" aria-hidden="true">
          <span class="show-label">Show</span>
          <span class="hide-label">Hide</span>
        </span>
      </summary>
      <div class="weekly-plan-body">
        <section class="weekly-story-card weekly-target-story">
          <div class="weekly-story-visual" aria-hidden="true">
            <span class="book-shape"></span>
            <span class="spark-shape"></span>
          </div>
          <div>
            <p class="eyebrow">Weekly Learning Target</p>
            <h3>${escapeHTML(storyTitle)}</h3>
            <p>${escapeHTML(plan.readingArticleSummary || `Read ${plan.readingTitle} and tell the main idea with details.`)}</p>
            <div class="weekly-story-actions">
              <a class="video-link-button" href="${weeklyVideoSearchUrl(plan)}" target="_blank" rel="noopener">Find read-aloud video</a>
            </div>
          </div>
        </section>
        <section class="weekly-knowledge-card">
          <p class="eyebrow">Knowledge Focus</p>
          <h3>This week's skills</h3>
          <div class="focus-chips">
            <span>${escapeHTML(plan.phonicsFocus || "phonics")}</span>
            <span>${escapeHTML(plan.grammarFocus || "grammar")}</span>
            <span>${escapeHTML(plan.uoiTopic || "topic vocabulary")}</span>
          </div>
        </section>
        <section class="weekly-vocab-card">
          <div class="weekly-vocab-head">
            <div>
              <p class="eyebrow">Words to Learn</p>
              <h3>Tap a word to learn it</h3>
            </div>
            <span>${vocabulary.length} words</span>
          </div>
          <div class="vocab-strip">
            ${(vocabularyGroups.length ? vocabularyGroups : [{ title: "Weekly Vocabulary", words: vocabulary }])
              .map((group) => `
                <article>
                  <h4>${escapeHTML(group.title)}</h4>
                  ${(group.words ?? []).map((item) => `
                    <button class="vocab-word-button" type="button" data-word="${escapeHTML(normalizeWord(item.word))}">${escapeHTML(item.word)}</button>
                  `).join("")}
                </article>
              `)
              .join("")}
          </div>
        </section>
      </div>
    </details>
  `;
}

function renderToday() {
  renderChildWeeklyPlan();
  document.querySelector("#childWeeklyPlan").hidden = false;
  document.querySelector("#missionGrid").hidden = false;
  document.querySelector(".epic-panel").hidden = false;
  document.querySelector("#practiceDetailPanel").hidden = true;

  const levels = getTodayAdventureLevels();
  const completedCount = levels.filter((level) => level.completed).length;
  const nextLevel = levels.find((level) => level.status === "next");
  document.querySelector("#missionGrid").innerHTML = `
    <section class="today-adventure-map" aria-label="Today Practice adventure map">
      <div class="adventure-map-head">
        <div>
          <p class="eyebrow">Today Practice</p>
          <h3>Treasure Map</h3>
          <p>${nextLevel ? `Next treasure stop: ${escapeHTML(nextLevel.title)}` : "Treasure found. Story star earned!"}</p>
        </div>
        <div class="adventure-score-badge" aria-label="${completedCount} of ${levels.length} levels complete">
          <strong>${completedCount}/${levels.length}</strong>
          <span>stars</span>
        </div>
      </div>
      ${treasureMapMarkup(levels)}
    </section>
  `;
}

const adventureMapLayout = {
  "read-warmup": { x: 16, y: 70, land: "lagoon" },
  speak: { x: 29, y: 43, land: "hill" },
  "word-power": { x: 42, y: 65, land: "forest" },
  listen: { x: 54, y: 34, land: "bridge" },
  sight: { x: 65, y: 56, land: "sign" },
  wordwork: { x: 77, y: 31, land: "tower" },
  "story-reward": { x: 86, y: 70, land: "chest" }
};

function treasureMapMarkup(levels) {
  return `
    <div class="treasure-map-board">
      <span class="map-deco map-cloud one" aria-hidden="true"></span>
      <span class="map-deco map-cloud two" aria-hidden="true"></span>
      <span class="map-deco map-river" aria-hidden="true"></span>
      <span class="map-deco map-hills" aria-hidden="true"></span>
      <span class="map-deco map-castle" aria-hidden="true"></span>
      <span class="map-deco map-cottage" aria-hidden="true"></span>
      <span class="map-deco map-grove" aria-hidden="true"></span>
      <span class="map-deco map-bridge-deco" aria-hidden="true"></span>
      <span class="map-deco map-treasure-cave" aria-hidden="true"></span>
      <span class="map-deco map-sun" aria-hidden="true"></span>
      <span class="map-deco map-water one" aria-hidden="true"></span>
      <span class="map-deco map-water two" aria-hidden="true"></span>
      <span class="map-deco map-palm" aria-hidden="true"></span>
      <span class="map-deco map-x" aria-hidden="true">X</span>
      <svg class="adventure-trail" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
        <path class="trail-shadow" d="M 16 70 C 18 58 22 46 29 43 S 36 60 42 65 S 48 42 54 34 S 60 47 65 56 S 72 42 77 31 S 84 48 86 70" />
        <path class="trail-line" d="M 16 70 C 18 58 22 46 29 43 S 36 60 42 65 S 48 42 54 34 S 60 47 65 56 S 72 42 77 31 S 84 48 86 70" />
      </svg>
      ${levels.map(treasureMapPin).join("")}
    </div>
  `;
}

function adventureLevelIcon(levelId) {
  const icon = {
    "read-warmup": "read",
    speak: "retell",
    "word-power": "word",
    listen: "spell",
    sight: "sight",
    wordwork: "grammar",
    "story-reward": "reward"
  }[levelId] ?? "star";
  return `<span class="map-icon map-icon-${icon}" aria-hidden="true"><span></span></span>`;
}

function adventureLevelSubtitle(levelId) {
  return {
    "read-warmup": "Listen, read, record",
    speak: "Tell it your way",
    "word-power": "Read today's words",
    listen: "Spell by sound",
    sight: "Find fast words",
    wordwork: "Build sentences",
    "story-reward": "Earned story time"
  }[levelId] ?? "Practice time";
}

function adventureLevelButtonLabel(level) {
  if (level.completed) return "Play again";
  if (!level.unlocked) return "Locked";
  return "Start Level";
}

function treasureMapPin(level) {
  const point = adventureMapLayout[level.id] ?? { x: 50, y: 50, land: "island" };
  return `
    <button
      class="map-level-pin ${level.status} landmark-${point.land}"
      type="button"
      data-start-practice="${escapeHTML(level.id)}"
      data-level-id="${escapeHTML(level.id)}"
      style="--map-x: ${point.x}%; --map-y: ${point.y}%;"
      aria-label="Level ${level.level}: ${escapeHTML(level.title)}. ${level.completed ? "Clear" : level.unlocked ? "Ready" : "Locked"}"
    >
      <span class="map-level-landmark" aria-hidden="true">${adventureLevelIcon(level.id)}</span>
      <span class="map-level-number">${level.level}</span>
      <span class="map-level-label">
        <strong>${escapeHTML(level.title)}</strong>
        <small>${level.completed ? "Clear" : level.unlocked ? "Start" : "Locked"}</small>
      </span>
    </button>
  `;
}

function practiceGroupProgress(centerIds) {
  const centers = centerIds.map((id) => missionCenters.find((item) => item.id === id)).filter(Boolean);
  const completed = centers.filter((center) => centerProgress(center).complete).length;
  const percent = centers.length ? Math.round((completed / centers.length) * 100) : 0;
  return {
    complete: completed === centers.length,
    percent,
    status: completed === centers.length ? "Done" : "Ready",
    label: `${completed}/${centers.length} parts complete`
  };
}

function renderCompactWeeklyHeader() {
  const plan = activePlan();
  return `
    <section class="kid-focus-strip">
      <div>
        <p class="eyebrow">This week</p>
        <h3>${escapeHTML(plan.readingArticleTitle || plan.readingTitle)}</h3>
      </div>
      <span>${escapeHTML((plan.masteryWords ?? []).length)} words</span>
      <span>Friday goal</span>
    </section>
  `;
}

function renderPracticeRoute() {
  const route = currentPracticeRoute();
  const panel = document.querySelector("#practiceRouteView");
  if (!panel) return;
  if (!route) {
    panel.innerHTML = "";
    return;
  }
  if (adventureLevelIds.includes(route.practiceId) && !isAdventureLevelUnlocked(route.practiceId)) {
    panel.innerHTML = lockedAdventurePageMarkup(route.practiceId);
    return;
  }
  if (route.practiceId === "read-warmup") {
    panel.innerHTML = readingWarmupPageMarkup();
    if (isSummerRazDate()) window.requestAnimationFrame(() => renderRazPdfViewer());
    return;
  }
  if (route.practiceId === "word-power") {
    panel.innerHTML = wordPowerPageMarkup();
    return;
  }
  if (route.practiceId === "story-reward") {
    panel.innerHTML = storyRewardPageMarkup();
    return;
  }
  if (route.practiceId === "story") {
    panel.innerHTML = storyPracticePageMarkup();
    return;
  }
  if (route.practiceId === "writing-lab") {
    panel.innerHTML = writingPracticePageMarkup();
    return;
  }
  const center = missionCenters.find((item) => item.id === route.practiceId) ?? missionCenters[0];
  panel.innerHTML = practicePageMarkup(center);
}

function lockedAdventurePageMarkup(levelId) {
  const level = adventureLevelDefinitions.find((item) => item.id === levelId);
  const previous = adventureLevelDefinitions[(level?.level ?? 1) - 2];
  return `
    <section class="wide-panel detail-card practice-page-card adventure-page-card locked-level-card">
      <button class="secondary-action" type="button" id="backToTodayButton">Back</button>
      <div class="practice-workspace">
        <h3>${escapeHTML(level?.title ?? "Adventure level")} is locked.</h3>
        <p>Finish ${escapeHTML(previous?.title ?? "the level before it")} first.</p>
      </div>
    </section>
  `;
}

function readingWarmupPageMarkup() {
  if (isSummerRazDate()) return summerRazReadingWarmupPageMarkup();
  const schoolReading = schoolReadingWarmupSource();
  const dayProgress = ensureAdventureDayProgress(activeProgress());
  const completed = isAdventureLevelCompleted("read-warmup");
  const skipped = Boolean(dayProgress.readingWarmupSkipped);
  const hasRecording = dayProgress.readingWarmupRecordingDone;
  const canComplete = hasRecording && dayProgress.readingWarmupReplayDone && !completed;
  const replayLabel = state.readingWarmupReplayActive ? "Stop replay" : "Replay";
  const skipStatus = readingWarmupSkipStatus();
  const skipHelper = skipStatus.required
    ? "Listen & Read is required today."
    : skipped
      ? "Skipped today. You can still read for practice."
      : skipStatus.canSkip
        ? `You can skip ${skipStatus.remaining} more Listen & Read day${skipStatus.remaining === 1 ? "" : "s"} this week.`
        : "No skips left this week. Please read today.";
  return `
    <section class="wide-panel detail-card practice-page-card story-page-card adventure-page-card reading-warmup-card">
      <button class="secondary-action" type="button" id="backToTodayButton">Back</button>
      <div class="practice-page-hero">
        <div class="story-page-art" aria-hidden="true">📖</div>
        <div>
          <p class="eyebrow">Level 1 · Listen & Read</p>
          <h2>${escapeHTML(schoolReading.title)}</h2>
          <p>${escapeHTML(schoolReading.summary)}</p>
        </div>
      </div>
      <div class="practice-workspace reading-workspace">
        <div class="warmup-reading-card">
          <p class="eyebrow">Read aloud warmup</p>
          <p>Open the read-aloud video, follow along with your paper book, then record yourself reading.</p>
          ${schoolReading.sourceLabel ? `<p>${escapeHTML(schoolReading.sourceLabel)}</p>` : ""}
          <a class="video-link-button" href="${escapeHTML(schoolReading.videoUrl)}" target="_blank" rel="noopener">Find read-aloud video</a>
        </div>
        <div class="retell-question-actions adventure-action-row">
          <button class="secondary-action" type="button" data-reading-warmup-record>${state.readingWarmupRecordingActive ? "Stop" : hasRecording ? "Record again" : "Record"}</button>
          <button class="secondary-action" type="button" data-reading-warmup-replay ${state.readingWarmupRecordingUrl ? "" : "disabled"}>${replayLabel}</button>
          <button class="primary-action" type="button" data-reading-warmup-complete ${canComplete ? "" : "disabled"}>${completed ? "Completed" : "I read it!"}</button>
          ${!skipStatus.required && !completed ? `<button class="secondary-action" type="button" data-reading-warmup-skip ${skipStatus.canSkip ? "" : "disabled"}>Skip today</button>` : ""}
        </div>
        ${state.readingWarmupReplayError ? `<p class="adventure-helper">${escapeHTML(state.readingWarmupReplayError)}</p>` : ""}
        <p class="adventure-helper">${completed ? skipped ? "Level 1 skipped today." : "Level 1 is complete." : hasRecording ? dayProgress.readingWarmupReplayDone ? "Replay checked. You can finish Level 1." : "Listen to your replay once, then finish Level 1." : "Record once, then replay your reading."}</p>
        <p class="adventure-helper">${escapeHTML(skipHelper)}</p>
      </div>
    </section>
  `;
}

function syncRazReadingAggregate(dayProgress = ensureAdventureDayProgress(activeProgress())) {
  const states = (dayProgress.razReadingAssignments ?? []).map((articleId) => dayProgress.razReadingArticles?.[articleId] ?? {});
  dayProgress.readingWarmupRecordingDone = states.length > 0 && states.every((article) => article.recordingDone);
  dayProgress.readingWarmupReplayDone = states.length > 0 && states.every((article) => article.replayDone);
  return dayProgress;
}

function razInventoryIsLow(date = todayKey()) {
  const nextDate = addDays(new Date(`${date}T00:00:00`), 1);
  return nextDate <= RAZ_SUMMER_END_DATE && summerRazAssignmentsForDate(nextDate).length < 2;
}

function summerRazReadingWarmupPageMarkup() {
  const dayProgress = syncRazReadingAggregate();
  const articles = razAssignmentsFromProgress(dayProgress);
  queueRazArticleAnalysis(articles);
  const isRecording = state.readingWarmupRecordingActive && state.mediaRecorder?.state === "recording";
  if (!articles.some((article) => article.id === state.razActiveArticleId)) state.razActiveArticleId = articles[0]?.id || "";
  const activeArticle = articles.find((article) => article.id === state.razActiveArticleId) || articles[0];
  const activeStatus = activeArticle ? dayProgress.razReadingArticles?.[activeArticle.id] ?? {} : {};
  const activeRecordingUrl = activeArticle ? state.readingWarmupRecordingUrls?.[activeArticle.id] : "";
  const completed = isAdventureLevelCompleted("read-warmup");
  const canComplete = dayProgress.readingWarmupRecordingDone && dayProgress.readingWarmupReplayDone && !completed;
  const skipStatus = readingWarmupSkipStatus();
  const finishedCount = articles.filter((article) => {
    const status = dayProgress.razReadingArticles?.[article.id] ?? {};
    return status.recordingDone && status.replayDone;
  }).length;
  const detailsStatus = activeArticle ? store.razArticleDetails?.[activeArticle.id]?.status : "";
  return `
    <section class="wide-panel detail-card practice-page-card adventure-page-card raz-reading-page">
      <button class="secondary-action" type="button" id="backToTodayButton" ${isRecording ? "disabled" : ""}>Back</button>
      <div class="detail-header raz-reading-header">
        <div>
          <p class="eyebrow">Level 1 · Summer Reading</p>
          <h2>Read two RAZ books aloud</h2>
          <p>${articles.length ? `${finishedCount}/${articles.length} books recorded and replayed` : "No RAZ books are available for today."}</p>
        </div>
      </div>
      <div class="practice-workspace raz-reading-workspace">
        ${articles.length ? `
          <div class="raz-article-tabs" role="tablist" aria-label="Today's RAZ books">
            ${articles.map((article) => {
              const status = dayProgress.razReadingArticles?.[article.id] ?? {};
              const done = status.recordingDone && status.replayDone;
              return `<button class="raz-article-tab ${article.id === activeArticle?.id ? "active" : ""} ${done ? "done" : ""}" type="button" role="tab" data-raz-article-select="${escapeHTML(article.id)}" ${isRecording ? "disabled" : ""}><span>Level ${escapeHTML(article.level)}</span><strong>${escapeHTML(article.title)}</strong>${done ? "<small>Done</small>" : ""}</button>`;
            }).join("")}
          </div>
          <div class="raz-recording-toolbar ${isRecording ? "recording" : ""}">
            <div class="retell-question-actions adventure-action-row">
              <button class="secondary-action" type="button" data-reading-warmup-record data-raz-article-id="${escapeHTML(activeArticle.id)}">${isRecording && state.readingWarmupActiveArticleId === activeArticle.id ? "Stop recording" : activeStatus.recordingDone ? "Record again" : "Start recording"}</button>
              <button class="secondary-action" type="button" data-reading-warmup-replay data-raz-article-id="${escapeHTML(activeArticle.id)}" ${activeRecordingUrl ? "" : "disabled"}>${state.readingWarmupReplayActive && state.readingWarmupReplayArticleId === activeArticle.id ? "Stop replay" : "Replay"}</button>
              <button class="primary-action" type="button" data-reading-warmup-complete ${canComplete ? "" : "disabled"}>${completed ? "Completed" : "Finish Level 1"}</button>
              ${!skipStatus.required && !completed ? `<button class="secondary-action" type="button" data-reading-warmup-skip ${skipStatus.canSkip ? "" : "disabled"}>Skip today</button>` : ""}
            </div>
            <p class="adventure-helper">${isRecording ? "Recording now. Keep reading and scroll through every page." : activeStatus.recordingDone ? activeStatus.replayDone ? "This book is done. Choose the other book." : "Replay this recording once." : "Start recording, then read and scroll through the book."}</p>
          </div>
          <div class="raz-pdf-panel">
            <div class="raz-pdf-heading">
              <div><p class="eyebrow">Level ${escapeHTML(activeArticle.level)}</p><h3>${escapeHTML(activeArticle.title)}</h3></div>
              <a class="secondary-action" href="${escapeHTML(encodeURI(activeArticle.path))}" target="_blank" rel="noopener">Full screen</a>
            </div>
            <div class="raz-pdf-viewer" data-raz-pdf-viewer data-article-id="${escapeHTML(activeArticle.id)}" data-pdf-url="${escapeHTML(activeArticle.path)}" aria-live="polite">
              <p class="raz-pdf-status" data-raz-pdf-status>Loading book...</p>
              <div class="raz-pdf-pages" data-raz-pdf-pages></div>
            </div>
          </div>
          ${state.readingWarmupReplayError ? `<p class="adventure-helper">${escapeHTML(state.readingWarmupReplayError)}</p>` : ""}
          ${detailsStatus === "failed" ? `<p class="adventure-helper">The PDF is ready to read. Level 2 will use general five-finger prompts because text details could not be extracted.</p>` : ""}
        ` : `
          <div class="retell-complete-panel"><h3>Add new RAZ books</h3><p>Ask a parent to add PDFs to material/H.PDF, material/I.PDF, or material/J.PDF, rebuild the RAZ manifest, and refresh.</p></div>
          ${!skipStatus.required ? `<button class="secondary-action" type="button" data-reading-warmup-skip ${skipStatus.canSkip ? "" : "disabled"}>Skip today</button>` : ""}
        `}
      </div>
      ${state.razInventoryPopup ? `<div class="raz-inventory-overlay" role="dialog" aria-modal="true" aria-labelledby="razInventoryTitle"><div class="raz-inventory-dialog"><h3 id="razInventoryTitle">Add new RAZ books</h3><p>The upcoming schedule needs more PDFs. Add files to the H, I, or J material folder, rebuild the manifest, and refresh.</p><button class="primary-action" type="button" data-close-raz-inventory>Continue</button></div></div>` : ""}
    </section>
  `;
}

function wordPowerPageMarkup() {
  const vocabulary = weeklyVocabulary();
  const status = getWordPowerStatus();
  const completedKeys = new Set(status.completedWords);
  const displayCount = Math.min(status.count, status.target);
  return `
    <section class="wide-panel detail-card practice-page-card adventure-page-card word-power-card">
      <button class="secondary-action" type="button" id="backToTodayButton">Back</button>
      <div class="detail-header">
        <div class="mission-visual sight" aria-hidden="true"><span>★</span></div>
        <div>
          <p class="eyebrow">Level 3 · Word Power</p>
          <h2>Words to Learn</h2>
          <div class="mini-progress detail-progress"><span style="width: ${status.target ? Math.min(100, Math.round((displayCount / status.target) * 100)) : 100}%"></span></div>
          <strong class="launcher-progress">${displayCount}/${status.target} words</strong>
        </div>
      </div>
      <div class="practice-workspace">
        <p class="adventure-helper">${vocabulary.length ? "Tap every word, hear it, and collect word power." : "No Words to Learn are set for today."}</p>
        <div class="word-tile-row">
          ${vocabulary.map((item) => {
            const key = normalizeWord(item.word);
            return `<button class="vocab-word-button ${completedKeys.has(key) ? "done" : ""}" type="button" data-word-power-word="${escapeHTML(key)}" data-word="${escapeHTML(key)}">${escapeHTML(item.word)}</button>`;
          }).join("")}
        </div>
        <button class="primary-action" type="button" data-complete-word-power ${status.ready ? "" : "disabled"}>${isAdventureLevelCompleted("word-power") ? "Words done!" : "Words done!"}</button>
      </div>
    </section>
  `;
}

function storyRewardPageMarkup() {
  const session = ensureStoryRewardQuizSession();
  const story = session?.story;
  const completed = isAdventureLevelCompleted("story-reward");
  if (!story) {
    return `
      <section class="wide-panel detail-card practice-page-card story-page-card adventure-page-card story-reward-card">
        <button class="secondary-action" type="button" id="backToTodayButton">Back</button>
        <div class="practice-page-hero">
          <div class="story-page-art" aria-hidden="true">📖</div>
          <div>
            <p class="eyebrow">Level 7 · Story Reward</p>
            <h2>Story audio is not ready yet.</h2>
            <p>Ask a parent to add this week's short reward audio.</p>
          </div>
        </div>
        <div class="practice-workspace reading-workspace">
          <p class="adventure-helper">Story Reward needs a project audio file under 5 minutes. No local generated story will be used here.</p>
        </div>
      </section>
    `;
  }
  const audioError = state.storyRewardAudioError;
  return `
    <section class="wide-panel detail-card practice-page-card story-page-card adventure-page-card story-reward-card">
      <button class="secondary-action" type="button" id="backToTodayButton">Back</button>
      <div class="practice-page-hero">
        <div class="story-page-art" aria-hidden="true">📖</div>
        <div>
          <p class="eyebrow">Level 7 · Story Reward</p>
          <h2>${escapeHTML(story.storyTitle)}</h2>
          <p>You earned story time. Enjoy the story one more time.</p>
        </div>
      </div>
      <div class="story-reward-flow">
        ${storyListenMarkup(story, "story-reward")}
        ${audioError ? `<p class="adventure-helper story-audio-error">${escapeHTML(audioError)}</p>` : ""}
        ${audioError ? "" : renderReadingQuizPractice({ mode: "story-reward", session })}
        ${completed ? `<p class="adventure-helper">Story reward complete.</p>` : ""}
      </div>
    </section>
  `;
}

function renderPracticeDetail() {
  const center = missionCenters.find((item) => item.id === state.activePracticeId) ?? missionCenters[0];
  document.querySelector("#practiceDetailPanel").innerHTML = practicePageMarkup(center);
}

function storyPracticePageMarkup() {
  const plan = activePlan();
  const story = currentListeningStory();
  const retell = buildWeeklyRetellPractice(plan);
  const title = retell.weeklyTitle;
  const storyProgress = practiceGroupProgress(["reading", "speak"]);
  const latestSpeech = (activeProgress().speakingAttempts ?? [])
    .filter((item) => item.date === todayKey() && item.source === "weekly-reading-retell")
    .slice(-1)[0];
  return `
    <section class="wide-panel detail-card practice-page-card story-page-card">
      <button class="secondary-action" type="button" id="backToTodayButton">Back</button>
      <div class="practice-page-hero">
        <div class="story-page-art" aria-hidden="true">📖</div>
        <div>
          <p class="eyebrow">Story Practice</p>
          <h2>${escapeHTML(title)}</h2>
          <p>${escapeHTML(plan.readingArticleSummary || "Listen, answer, and retell this week's story.")}</p>
          <div class="mini-progress detail-progress"><span style="width: ${storyProgress.percent}%"></span></div>
          <strong class="launcher-progress">${escapeHTML(storyProgress.label)}</strong>
        </div>
      </div>
      <div class="story-practice-grid">
        <article class="story-task-card listen-quiz-card">
          <span class="task-number">1</span>
          <p class="eyebrow">Listen & Quiz</p>
          <h3>Hear Story: ${escapeHTML(story.storyTitle)}</h3>
          <img class="story-cover-card" src="${storyCoverSrc(story)}" alt="${escapeHTML(story.cover?.alt ?? story.storyTitle)}" width="240" height="160" />
          ${storyListenMarkup(story, "story-preview")}
          <button class="primary-action" type="button" data-start-practice="reading">Start Quiz</button>
        </article>
        <article class="story-task-card speak-card">
          <span class="task-number">2</span>
          <p class="eyebrow">Speak</p>
          <h3>Answer the following questions according to ${escapeHTML(title)}</h3>
          <div class="feedback-preview">${latestSpeech?.feedback ? escapeHTML(latestSpeech.feedback) : "Feedback will appear here after speaking practice."}</div>
          <button class="primary-action" type="button" data-start-practice="speak">Start Speaking</button>
        </article>
      </div>
    </section>
  `;
}

function writingPracticePageMarkup() {
  const plan = activePlan();
  const writingProgress = practiceGroupProgress(["listen", "sight", "wordwork"]);
  const spelling = missionCenters.find((item) => item.id === "listen");
  const sight = missionCenters.find((item) => item.id === "sight");
  const grammar = missionCenters.find((item) => item.id === "wordwork");
  return `
    <section class="wide-panel detail-card practice-page-card writing-page-card">
      <button class="secondary-action" type="button" id="backToTodayButton">Back</button>
      <div class="practice-page-hero">
        <div class="writing-page-art" aria-hidden="true">✎</div>
        <div>
          <p class="eyebrow">Writing Practice</p>
          <h2>Words, sight words, and grammar</h2>
          <p>${escapeHTML(plan.grammarFocus || "Practice this week's grammar focus with short questions.")}</p>
          <div class="mini-progress detail-progress"><span style="width: ${writingProgress.percent}%"></span></div>
          <strong class="launcher-progress">${escapeHTML(writingProgress.label)}</strong>
        </div>
      </div>
      <div class="writing-segment-row" aria-label="Writing practice sections">
        <span>Spelling Words</span>
        <span>Sight Words</span>
        <span>Grammar Practice</span>
      </div>
      <div class="writing-task-grid">
        ${writingTaskCard(spelling, "Spelling Words", "Hear each weekly word, then spell it.", "listen")}
        ${writingTaskCard(sight, "Sight Words", "Read weekly sight words quickly and carefully.", "sight")}
        ${writingTaskCard(grammar, "Grammar Practice", plan.grammarFocus || "Answer short grammar questions from this week.", "wordwork")}
      </div>
    </section>
  `;
}

function writingTaskCard(center, title, description, practiceId) {
  const progress = center ? centerProgress(center) : { percent: 0, status: "Ready", label: "0/1 complete", complete: false };
  return `
    <article class="writing-task-card ${progress.complete ? "done" : ""}">
      <span class="domain-chip">${progress.status}</span>
      <h3>${escapeHTML(title)}</h3>
      <p>${escapeHTML(description)}</p>
      <div class="mini-progress"><span style="width: ${progress.percent}%"></span></div>
      <strong class="launcher-progress">${escapeHTML(progress.label)}</strong>
      <button class="primary-action" type="button" data-start-practice="${practiceId}">Start</button>
    </article>
  `;
}

function practicePageMarkup(center) {
  const plan = activePlan();
  const word = center.id === "sight" ? plan.sightWords[0] : plan.masteryWords[0];
  const progress = centerProgress(center);
  const detailByType = {
    spelling: renderListenPractice(),
    flashcards: renderSightPractice(),
    grammar: renderGrammarPractice(),
    speaking: renderSpeakingPractice(),
    reading: renderReadingQuizPractice(),
    writing: `
      <div class="practice-workspace writing-workspace">
        <h3>Use a weekly word in a complete sentence.</h3>
        <p>Try: ${escapeHTML(word ?? "weekly word")}</p>
        <textarea placeholder="Write one complete sentence."></textarea>
        <button class="primary-action" type="button" data-practice-action="writing">Submit sentence</button>
      </div>
    `,
    review: `
      <div class="practice-workspace review-workspace">
        <h3>Review due words</h3>
        <div class="word-tile-row">${(dueReviewWords().map((item) => item.word).slice(0, 8).length ? dueReviewWords().map((item) => item.word).slice(0, 8) : plan.masteryWords.slice(0, 6)).map((item) => `<button type="button" data-practice-action="warmup">${escapeHTML(item)}</button>`).join("")}</div>
      </div>
    `
  };
  return `
    <section class="wide-panel detail-card practice-page-card ${center.id === "listen" ? "spelling-page-card" : ""}">
      <button class="secondary-action" type="button" id="backToTodayButton">Back</button>
      <div class="detail-header">
        <div class="mission-visual ${center.visual}" aria-hidden="true"><span>${center.visualIcon}</span></div>
        <div>
          <p class="eyebrow">${center.domain}</p>
          <h2>${center.title}</h2>
          <div class="mini-progress detail-progress"><span style="width: ${progress.percent}%"></span></div>
          <div class="practice-route-meta">
            <strong class="launcher-progress">${escapeHTML(progress.label)}</strong>
            <span class="timer-chip" id="practiceTimerChip">${formatDuration(elapsedPracticeSeconds(center.id))}</span>
          </div>
        </div>
      </div>
      ${detailByType[center.practiceType] ?? detailByType.review}
    </section>
  `;
}

function renderQuestionChoices(question, attributeName) {
  return question.choices
    .map((choice, index) => `<button type="button" data-${attributeName}="${index}">${escapeHTML(choice)}</button>`)
    .join("");
}

function renderGrammarQuestionChoices(question, feedback) {
  return question.choices
    .map((choice, index) => {
      const classes = ["quiz-choice"];
      if (feedback) {
        if (index === feedback.answerIndex) classes.push("correct-choice");
        if (index === feedback.selectedIndex && !feedback.correct) classes.push("wrong-choice");
        if (index === feedback.selectedIndex) classes.push("selected-choice");
      }
      return `<button class="${classes.join(" ")}" type="button" data-grammar-choice="${index}" ${feedback ? "disabled" : ""}>${escapeHTML(choice)}</button>`;
    })
    .join("");
}

function storyCoverSrc(story) {
  return story.coverUrl || story.cover?.imageDataUri || "";
}

function storyListenMarkup(story, key = "story-preview") {
  if ((story.sourceType === "network" || story.url) && hasPlayableAudio(story)) {
    const rewardAudioAttrs = key === "story-reward"
      ? ` data-story-reward-audio data-declared-duration="${escapeHTML(String(story.durationSeconds || ""))}"`
      : "";
    return `
      <div class="audio-story-shell network-audio-shell">
        <div class="network-audio-player">
          <audio controls preload="metadata" src="${escapeHTML(story.audioUrl)}"${rewardAudioAttrs}></audio>
        </div>
        <div class="network-audio-meta">
          <strong>${escapeHTML(story.sourceName || "Online story")}</strong>
          <span>${escapeHTML(story.skillMatch || "Open the story page to listen or read.")}</span>
          ${story.durationSeconds ? `<span>${Math.ceil(Number(story.durationSeconds) / 60)} min or less</span>` : ""}
        </div>
        <a class="secondary-action" href="${escapeHTML(story.url)}" target="_blank" rel="noopener noreferrer">Source</a>
      </div>
    `;
  }
  return `
    <div class="audio-story-shell">
      <button class="word-audio-button" type="button" data-audio-target="${escapeHTML(`${story.storyTitle}. ${story.storyText}`)}" data-audio-key="${escapeHTML(key)}">▶</button>
      <div>
        <strong>${escapeHTML(story.storyTitle)}</strong>
        <span>1-2 minute story for this week's target</span>
      </div>
    </div>
  `;
}

function storyReplayMarkup(story, key) {
  if ((story.sourceType === "network" || story.url) && hasPlayableAudio(story)) {
    return `<audio class="story-replay-button" controls preload="metadata" src="${escapeHTML(story.audioUrl)}"></audio>`;
  }
  return `<button class="secondary-action story-replay-button" type="button" data-audio-target="${escapeHTML(`${story.storyTitle}. ${story.storyText}`)}" data-audio-key="${escapeHTML(key)}">Replay Story</button>`;
}

function renderReadingQuestionChoices(question, feedback) {
  return question.choices
    .map((choice, index) => {
      const classes = ["quiz-choice"];
      if (feedback) {
        if (index === feedback.answerIndex) classes.push("correct-choice");
        if (index === feedback.selectedIndex && !feedback.correct) classes.push("wrong-choice");
        if (index === feedback.selectedIndex) classes.push("selected-choice");
      }
      return `<button class="${classes.join(" ")}" type="button" data-reading-quiz-choice="${index}" ${feedback ? "disabled" : ""}>${escapeHTML(choice)}</button>`;
    })
    .join("");
}

function renderReadingQuizPractice({ mode = "reading", session = null } = {}) {
  const activeSession = session || currentListeningStorySession();
  if (!activeSession) {
    return `
      <div class="practice-workspace reading-workspace">
        <h3>Story audio is not ready yet.</h3>
        <p>No quiz can start until this week's reward audio is available.</p>
      </div>
    `;
  }
  const story = activeSession.story ?? currentListeningStory();
  const questions = story.questions;
  const isStoryReward = mode === "story-reward";
  const answeredCount = activeSession.answers?.length ?? 0;
  const currentIndex = Math.min(activeSession.currentIndex ?? 0, Math.max(questions.length - 1, 0));
  const question = questions[currentIndex];
  const feedback = activeSession.pendingFeedback;
  const percent = questions.length ? Math.round((answeredCount / questions.length) * 100) : 0;
  if (activeSession.completed) {
    return `
      <div class="practice-workspace reading-workspace">
        <h3>${isStoryReward ? "Story quiz complete!" : "Quiz complete!"}</h3>
        <p>You got ${activeSession.score} of ${questions.length} questions correct.</p>
        ${
          isStoryReward
            ? `<button class="primary-action" type="button" data-finish-story-reward>${isAdventureLevelCompleted("story-reward") ? "Story reward complete" : "Finish story reward"}</button>`
            : `<button class="primary-action" type="button" data-start-practice="story">Back to Story</button>`
        }
      </div>
    `;
  }
  return `
    <div class="practice-workspace reading-workspace">
      <div class="quiz-story-head">
        <img class="story-cover-thumb" src="${storyCoverSrc(story)}" alt="${escapeHTML(story.cover?.alt ?? story.storyTitle)}" width="120" height="80" />
        <div>
          <p class="eyebrow">Story Quiz</p>
          <h3>${escapeHTML(story.storyTitle)}</h3>
        </div>
      </div>
      <div class="mini-progress detail-progress" aria-label="Quiz progress"><span style="width: ${percent}%"></span></div>
      <div class="quiz-question-head">
        <div>
          <p class="eyebrow">Question ${currentIndex + 1} of ${questions.length}</p>
          <h3>${escapeHTML(question.prompt)}</h3>
        </div>
        ${storyReplayMarkup(story, `quiz-story-replay-${currentIndex}`)}
      </div>
      <div class="word-tile-row">
        ${renderReadingQuestionChoices(question, feedback)}
      </div>
      ${
        feedback
          ? `<div class="quiz-feedback ${feedback.correct ? "correct" : "wrong"}" role="status">
              <strong>${feedback.correct ? "Correct" : "Try again"}</strong>
              <span>${feedback.correct ? "Nice listening." : `The answer is ${escapeHTML(question.answer)}.`}</span>
            </div>
            <button class="primary-action" type="button" data-reading-quiz-next>${currentIndex + 1 >= questions.length ? "Finish Quiz" : "Next Question"}</button>`
          : ""
      }
    </div>
  `;
}

function renderSpeakingPractice() {
  const retell = buildWeeklyRetellPractice();
  if (state.retellSessionKey !== retell.sourceKey) {
    state.retellSessionKey = retell.sourceKey;
    state.retellParentChecked = retellParentChecksFromProgress(retell);
    state.speakingRecordingUrls = {};
    state.speakingReplayErrors = {};
  }
  if (retell.sourceType === "raz-five-finger") return renderSummerRazRetellPractice(retell);
  if (retell.missingStoryEvidence) {
    const missingCopy = retell.missingPoemEvidence
      ? {
          body: "Poem details are needed for today's Level 2 response. Ask a parent to add the poem evidence in Learning Materials.",
          path: "Parent path: Learning Materials → Review weekly plan → Poem details for Level 2 → Use for this week."
        }
      : {
          body: "Story pages are needed for today's retell. Ask a parent to upload the HMH story pages in Learning Materials.",
          path: "Parent path: Learning Materials → Material kind: Main Story Pages → Upload → Create weekly plan → review story details."
        };
    return `
      <div class="practice-workspace speaking-workspace">
        <h3>Retell ${escapeHTML(retell.weeklyTitle)}</h3>
        <div class="retell-complete-panel">
          <p>${escapeHTML(missingCopy.body)}</p>
          <p>${escapeHTML(missingCopy.path)}</p>
          <button class="secondary-action" type="button" id="backToTodayButton">Back</button>
        </div>
      </div>
    `;
  }
  const promptItems = retell.prompts.map((prompt) => normalizeRetellPrompt(prompt, retell));
  const latestSpeech = (activeProgress().speakingAttempts ?? [])
    .filter((item) => item.date === todayKey() && item.source === "weekly-reading-retell")
    .slice(-1)[0];
  state.speakingTarget = promptItems[0]?.prompt ?? "";
  const allPromptsChecked = promptItems.length > 0 && promptItems.every((_, index) => state.retellParentChecked?.[index]);
  return `
    <div class="practice-workspace speaking-workspace">
      <h3>Retell ${escapeHTML(retell.weeklyTitle)}</h3>
      <div class="retell-prompts">
        ${promptItems.map((prompt, index) => {
          const hasRecording = Boolean(state.speakingRecordingUrls?.[index]);
          const isRecording = state.activeSpeakingRecordingIndex === index && state.mediaRecorder?.state === "recording";
          const checked = Boolean(state.retellParentChecked?.[index]);
          const replayError = state.speakingReplayErrors?.[index] || "";
          return `
          <article class="retell-question-card">
            <p>${escapeHTML(prompt.prompt)}</p>
            <div class="retell-support">
              <p><strong>Try:</strong> ${prompt.hints.map(escapeHTML).join(" ")}</p>
              <p><strong>Words:</strong> ${prompt.keywords.map(escapeHTML).join(", ")}</p>
              ${checked ? `
                <p><strong>Reference:</strong> ${escapeHTML(prompt.referenceAnswer)}</p>
                <p><strong>Stronger:</strong> ${escapeHTML(prompt.strongerAnswer)}</p>
              ` : ""}
            </div>
            <div class="retell-question-actions">
              <button class="secondary-action" type="button" data-record-prompt-index="${index}">${isRecording ? "Stop" : hasRecording ? "Record again" : "Record"}</button>
              <button class="secondary-action" type="button" data-retell-replay="${index}" ${hasRecording ? "" : "disabled"}>Replay</button>
              <button class="primary-action" type="button" data-retell-parent-check="${index}" ${checked ? "disabled" : ""}>${checked ? "Checked with parents" : "Check answer with your parents"}</button>
            </div>
            ${replayError ? `<p class="adventure-helper">${escapeHTML(replayError)}</p>` : ""}
          </article>
        `;
        }).join("")}
      </div>
      ${allPromptsChecked ? `
        <div class="retell-complete-panel">
          <button class="primary-action" type="button" data-complete-retell-practice>Complete retell</button>
        </div>
      ` : ""}
      <div class="feedback-preview">${latestSpeech?.feedback ? escapeHTML(latestSpeech.feedback) : "Record, replay, then check with your parents."}</div>
    </div>
  `;
}

function renderSummerRazRetellPractice(retell) {
  const dayProgress = ensureAdventureDayProgress(activeProgress());
  const progress = dayProgress.razRetell;
  if (retell.missingStoryEvidence) {
    return `<div class="practice-workspace speaking-workspace"><h3>Five Finger Retell</h3><div class="retell-complete-panel"><p>No RAZ book is available today. Ask a parent to add new PDFs and rebuild the RAZ manifest.</p><button class="secondary-action" type="button" id="backToTodayButton">Back</button></div></div>`;
  }
  if (retell.needsArticleSelection) {
    return `
      <div class="practice-workspace speaking-workspace raz-retell-workspace">
        <h3>Choose one book to retell</h3>
        <div class="raz-retell-choices">
          ${retell.articles.map((article) => `<button class="raz-retell-choice" type="button" data-raz-retell-select="${escapeHTML(article.id)}"><span>Level ${escapeHTML(article.level)}</span><strong>${escapeHTML(article.title)}</strong></button>`).join("")}
        </div>
      </div>`;
  }
  const hasRecording = Boolean(state.speakingRecordingUrls?.[0]) || progress.recordingDone;
  const hasReplayUrl = Boolean(state.speakingRecordingUrls?.[0]);
  const isRecording = state.activeSpeakingRecordingIndex === 0 && state.mediaRecorder?.state === "recording";
  const analysisPending = Boolean(state.razArticleAnalysisPending[retell.selectedArticle?.id]);
  const detailsFailed = store.razArticleDetails?.[retell.selectedArticle?.id]?.status === "failed";
  const readyToCheck = progress.recordingDone && progress.replayDone;
  const readyToComplete = readyToCheck && progress.parentChecked;
  return `
    <div class="practice-workspace speaking-workspace raz-retell-workspace">
      <div class="raz-retell-title-row"><div><p class="eyebrow">Level 2 · Five Finger Retell</p><h3>${escapeHTML(retell.weeklyTitle)}</h3></div><button class="secondary-action" type="button" data-change-raz-retell-book ${isRecording ? "disabled" : ""}>Change book</button></div>
      <div class="five-finger-grid">
        ${retell.fingers.map((finger) => `<article class="five-finger-item"><strong>${escapeHTML(finger.label)}</strong><p>${escapeHTML(finger.prompt)}</p><small>${escapeHTML(finger.support)}</small></article>`).join("")}
      </div>
      <article class="retell-question-card raz-combined-retell">
        <p><strong>One recording:</strong> ${escapeHTML(retell.prompts[0].prompt)}</p>
        <div class="retell-question-actions">
          <button class="secondary-action" type="button" data-record-prompt-index="0">${isRecording ? "Stop" : hasRecording ? "Record again" : "Record retell"}</button>
          <button class="secondary-action" type="button" data-retell-replay="0" ${hasReplayUrl ? "" : "disabled"}>Replay</button>
          <button class="primary-action" type="button" data-retell-parent-check="0" ${readyToCheck && !progress.parentChecked ? "" : "disabled"}>${progress.parentChecked ? "Checked with parents" : "Parent check"}</button>
        </div>
        ${state.speakingReplayErrors?.[0] ? `<p class="adventure-helper">${escapeHTML(state.speakingReplayErrors[0])}</p>` : ""}
        <p class="adventure-helper">${progress.parentChecked ? "Parent check complete." : progress.replayDone ? "Now check the retell with a parent." : progress.recordingDone ? "Replay the whole retell once." : "Use the five prompts from top to bottom."}</p>
      </article>
      ${analysisPending ? `<p class="adventure-helper">Reading the PDF details for stronger prompts...</p>` : ""}
      ${detailsFailed ? `<p class="adventure-helper">General prompts are shown because the PDF text could not be extracted.</p>` : ""}
      ${readyToComplete ? `<div class="retell-complete-panel"><button class="primary-action" type="button" data-complete-retell-practice>Complete retell</button></div>` : ""}
    </div>`;
}

function renderGrammarPractice() {
  const session = ensureGrammarSession();
  const questions = session.questions ?? grammarQuestions();
  const answeredCount = session.answers?.length ?? 0;
  const currentIndex = Math.min(session.currentIndex ?? 0, Math.max(questions.length - 1, 0));
  const question = questions[currentIndex];
  const feedback = session.pendingFeedback;
  const percent = questions.length ? Math.round((answeredCount / questions.length) * 100) : 0;
  if (session.completed) {
    return `
      <div class="practice-workspace sort-workspace">
        <h3>Grammar complete!</h3>
        <p>You got ${session.score} of ${questions.length} questions correct.</p>
        <button class="primary-action reward-ready-action" type="button" data-grammar-reward-ready>🎁 Get your story reward</button>
      </div>
    `;
  }
  return `
    <div class="practice-workspace sort-workspace">
      <h3>${escapeHTML(activePlan().grammarFocus || "Grammar Practice")}</h3>
      <div class="mini-progress detail-progress" aria-label="Grammar progress"><span style="width: ${percent}%"></span></div>
      <p class="eyebrow">Question ${currentIndex + 1} of ${questions.length}</p>
      <h3>${escapeHTML(question.prompt)}</h3>
      <div class="word-tile-row">
        ${renderGrammarQuestionChoices(question, feedback)}
      </div>
      ${feedback ? `<div class="quiz-feedback ${feedback.correct ? "correct" : "wrong"}" role="status">
        ${feedback.correct ? "Correct!" : `Try again next time. Answer: ${escapeHTML(question.answer)}`}
      </div>` : ""}
    </div>
  `;
}

function renderListenPractice() {
  const session = ensureSpellingSession();
  const words = session.words ?? [];
  const currentIndex = Math.min(session.currentIndex ?? 0, Math.max(words.length - 1, 0));
  const currentWord = words[currentIndex] ?? "";
  const completed = spellingSessionCompletedCount(session);
  const progressPercent = words.length ? Math.round((completed / words.length) * 100) : 0;
  const attemptsUsed = session.attemptsForCurrent ?? 0;
  const retryChancesLeft = attemptsUsed
    ? Math.max(MAX_SPELLING_ATTEMPTS - attemptsUsed, 0)
    : MAX_SPELLING_RETRY_CHANCES;
  const overlay = state.spellingFeedbackOverlay;
  const wordPulseClass = state.spellingWordPulse ? "word-pulse" : "";
  const answer = session.currentAnswer ?? "";
  if (!words.length) {
    return `
      <div class="practice-workspace spelling-workspace playful-spelling">
        <h3>Words are coming soon.</h3>
        <p>Add weekly words in the Parent weekly plan, then come back to spell.</p>
      </div>
    `;
  }
  if (spellingSessionIsComplete(session)) {
    return `
      <div class="practice-workspace spelling-workspace playful-spelling spelling-complete">
        <div class="listen-orb" aria-hidden="true">★</div>
        <h3>Spelling center complete!</h3>
        <p>You checked ${words.length} words today.</p>
        <button class="primary-action" type="button" id="backToTodayButton">Back</button>
      </div>
    `;
  }
  return `
    <div class="practice-workspace spelling-workspace playful-spelling">
      <div class="spelling-sky" aria-hidden="true">
        <span class="spelling-cloud one"></span>
        <span class="spelling-cloud two"></span>
        <span class="letter-block block-a">a</span>
        <span class="letter-block block-b">b</span>
        <span class="letter-block block-c">c</span>
      </div>
      <div class="spelling-stage">
        <div class="listen-orb" aria-hidden="true">🎧</div>
        <div>
          <p class="eyebrow spelling-word-count ${wordPulseClass}">Word ${completed + 1} of ${words.length}</p>
          <h3>Listen first. Then spell.</h3>
          <div class="mini-progress spelling-progress" aria-label="Spelling progress">
            <span style="width: ${progressPercent}%"></span>
          </div>
        </div>
      </div>
      <div class="spelling-control-row">
        <button class="primary-action hear-word-action" type="button" data-spelling-audio="${escapeHTML(currentWord)}" data-audio-key="listen-word-${currentIndex}">Hear word</button>
        <span class="attempt-chip">${retryChancesLeft} retries left</span>
      </div>
      <form class="spelling-answer-form" id="spellingAnswerForm" autocomplete="off">
        <div class="spelling-answer-panel">
          <label class="spelling-answer-label" for="spellingAnswerInput">Spell the word</label>
          <input
            class="spelling-answer-display ${answer ? "" : "empty"}"
            id="spellingAnswerInput"
            name="spellingAnswer"
            type="text"
            value="${escapeHTML(answer)}"
            placeholder="Type the word"
            autocomplete="off"
            autocorrect="off"
            autocapitalize="none"
            spellcheck="false"
            inputmode="text"
            pattern="[A-Za-z']*"
            aria-label="Your spelling answer"
          />
        </div>
        <button class="primary-action" type="submit" id="checkSpellingButton" data-spelling-word="${escapeHTML(currentWord)}">Check</button>
      </form>
      ${overlay ? renderSpellingFeedbackOverlay(overlay) : ""}
    </div>
  `;
}

function renderSpellingFeedbackOverlay(overlay) {
  return `
    <div class="spelling-feedback-overlay ${overlay.type}" role="status" aria-live="assertive">
      <div class="feedback-card feedback-pop">
        <div class="feedback-icon ${overlay.icon}" aria-hidden="true"></div>
        <strong>${escapeHTML(overlay.title)}</strong>
        <span>${escapeHTML(overlay.message)}</span>
        ${overlay.word ? `<span class="feedback-spelling-word">Correct spelling: <b>${escapeHTML(displayWord(overlay.word))}</b></span>` : ""}
        ${overlay.type === "saved" ? `<button class="primary-action" type="button" data-close-spelling-feedback>Next word</button>` : ""}
      </div>
    </div>
  `;
}

function currentSightSessionState() {
  const session = ensureSightSession();
  const activities = session.activities ?? sightActivitiesForWords(session.words ?? []);
  const index = Math.min(session.currentIndex ?? 0, Math.max(activities.length - 1, 0));
  const activity = activities[index] ?? null;
  const word = activity?.word ?? "";
  return { session, activities, index, activity, word, mode: activity?.type ?? "hear-find" };
}

function renderSightPractice() {
  const { session, activities, index, activity, word, mode } = currentSightSessionState();
  const completed = Math.min(session.completedActivities?.length ?? session.completedWords?.length ?? 0, activities.length);
  const progressPercent = activities.length ? Math.round((completed / activities.length) * 100) : 0;
  const overlay = state.sightFeedbackOverlay;
  if (!activities.length) {
    return `<div class="practice-workspace sight-workspace"><h3>Words are coming soon.</h3><p>Add high-frequency words in the weekly plan.</p></div>`;
  }
  if (completed >= activities.length) {
    return `
      <div class="practice-workspace sight-workspace sight-complete">
        <div class="sight-badge" aria-hidden="true">★</div>
        <h3>Sight word center complete!</h3>
        <p>You practiced ${activities.length} quick word activities today.</p>
        <button class="primary-action" type="button" id="backToTodayButton">Back</button>
      </div>
    `;
  }
  const content = {
    "hear-find": renderSightHearFind(word, index),
    "sentence-cloze": renderSightSentence(word),
    "word-build": renderSightBuild(word),
    "mini-passage-hunt": renderSightHunt(word),
    "sentence-match": renderSightSentenceMatch(activity, index)
  }[mode];
  return `
    <div class="practice-workspace sight-workspace">
      <div class="sight-game-head">
        <div class="sight-badge" aria-hidden="true">★</div>
        <div>
          <p class="eyebrow">Question ${completed + 1} of ${activities.length}</p>
          <h3>${sightModeTitle(mode)}</h3>
          <div class="mini-progress sight-progress"><span style="width: ${progressPercent}%"></span></div>
        </div>
      </div>
      ${content}
      ${overlay ? renderSpellingFeedbackOverlay(overlay) : ""}
    </div>
  `;
}

function sightModeTitle(mode) {
  return {
    "hear-find": "Hear it. Find it.",
    "sentence-cloze": "Choose the word in the sentence.",
    "word-build": "Build the word.",
    "mini-passage-hunt": "Find the word fast.",
    "sentence-match": "Choose the sentence that makes sense."
  }[mode] ?? "Practice the word.";
}

function sightWordChunks(word) {
  const text = displayWord(word).toLowerCase();
  const patterns = ["igh", "ear", "air", "are", "oo", "ee", "ea", "ai", "ay", "oa", "ow", "ou", "oi", "oy", "ar", "er", "ir", "or", "ur", "sh", "ch", "th", "wh", "ck", "ng", "le"];
  const chunks = [];
  let index = 0;
  while (index < text.length) {
    const pattern = patterns.find((item) => text.slice(index).startsWith(item));
    if (pattern) {
      chunks.push(pattern);
      index += pattern.length;
    } else {
      chunks.push(text[index]);
      index += 1;
    }
  }
  return chunks.filter(Boolean);
}

function sightFocusPartIndex(chunks) {
  const multi = chunks.findIndex((part) => part.length > 1 && part !== "'");
  if (multi >= 0) return multi;
  const vowel = chunks.findIndex((part) => /[aeiou]/.test(part));
  if (vowel >= 0) return vowel;
  return Math.max(0, Math.floor(chunks.length / 2));
}

function sightPartChoices(answer, seed) {
  const pool = ["a", "e", "i", "o", "u", "oo", "ee", "ea", "ai", "ay", "oa", "ow", "ou", "ar", "er", "or", "th", "sh", "ch", "ck", "ng", "le"];
  const choices = uniqueWords([answer, ...seededShuffle(pool.filter((item) => item !== answer), `${state.activeLearnerId}:${todayKey()}:part:${seed}`).slice(0, 3)]);
  return seededShuffle(choices.slice(0, 4), `${state.activeLearnerId}:${todayKey()}:part-order:${seed}`);
}

function sightChoices(word, seed) {
  const key = normalizeWord(word);
  const distractors = seededShuffle(sightChoicePool(word).filter((item) => normalizeWord(item) !== key), `${state.activeLearnerId}:${todayKey()}:choices:${seed}`).slice(0, 3);
  return seededShuffle(uniqueWords([word, ...distractors]).slice(0, 4), `${state.activeLearnerId}:${todayKey()}:choice-order:${seed}`);
}

function renderSightHearFind(word, index) {
  return `
    <div class="sight-game-panel">
      ${audioButton("Hear word", word, `sight-hear-${index}`, "primary-action hear-word-action")}
      <p class="sight-helper">Tap the word you hear.</p>
      <div class="sight-choice-grid">
        ${sightChoices(word, index).map((choice) => `<button class="sight-choice" type="button" data-sight-choice="${escapeHTML(normalizeWord(choice))}">${escapeHTML(displayWord(choice))}</button>`).join("")}
      </div>
    </div>
  `;
}

function renderSightFill(word) {
  const chunks = sightWordChunks(word);
  const missingIndex = sightFocusPartIndex(chunks);
  const answer = chunks[missingIndex] ?? "";
  return `
    <div class="sight-game-panel">
      ${audioButton("Hear word", word, `sight-fill-${normalizeWord(word)}`, "primary-action hear-word-action")}
      <p class="sight-helper">Choose the missing sound part.</p>
      <div class="missing-word" aria-label="Word with a missing part">
        ${chunks.map((part, index) => `<span class="${index === missingIndex ? "blank-part" : ""}">${index === missingIndex ? "" : escapeHTML(part)}</span>`).join("")}
      </div>
      <div class="sight-choice-grid compact-choices">
        ${sightPartChoices(answer, word).map((choice) => `<button class="sight-choice sight-fill-choice" type="button" data-sight-fill="${escapeHTML(choice)}">${escapeHTML(choice)}</button>`).join("")}
      </div>
    </div>
  `;
}

function renderSightBuild(word) {
  const session = ensureSightSession();
  const activity = session.activities?.[session.currentIndex ?? 0];
  const key = normalizeWord(word);
  if (!session.buildState || session.buildState.word !== key) {
    session.buildState = { word: key, parts: [], usedIndexes: [] };
  }
  const targetParts = sightWordChunks(word);
  const tiles = seededShuffle(targetParts, `${state.activeLearnerId}:${todayKey()}:build:${word}:${activity?.id ?? ""}`);
  const built = session.buildState.parts ?? session.buildState.letters ?? [];
  const usedIndexes = session.buildState.usedIndexes ?? [];
  return `
    <div class="sight-game-panel">
      ${audioButton("Hear word", word, `sight-build-${normalizeWord(word)}`, "primary-action hear-word-action")}
      <p class="sight-helper">Tap the parts in order.</p>
      <div class="built-word-slots" aria-label="Built word">
        ${targetParts.map((_, index) => `<span>${escapeHTML(built[index] ?? "")}</span>`).join("")}
      </div>
      <div class="letter-tile-row">
        ${tiles.map((part, index) => `<button type="button" data-sight-build-letter="${index}" ${usedIndexes.includes(index) ? "disabled" : ""}>${escapeHTML(part)}</button>`).join("")}
      </div>
      <button class="secondary-action" type="button" data-sight-build-clear>Clear</button>
    </div>
  `;
}

function renderSightSentence(word) {
  const sentenceText = meaningfulSightSentence(word);
  const sentence = sentenceText.replace(new RegExp(`\\b${displayWord(word)}\\b`, "i"), "____");
  return `
    <div class="sight-game-panel">
      ${audioButton("Hear sentence", sentenceText, `sight-sentence-${normalizeWord(word)}`, "primary-action hear-word-action")}
      <p class="sentence-strip">${escapeHTML(sentence.includes("____") ? sentence : meaningfulSightSentenceBlank(word))}</p>
      <div class="sight-choice-grid">
        ${sightChoices(word, `sentence-${word}`).map((choice) => `<button class="sight-choice" type="button" data-sight-choice="${escapeHTML(normalizeWord(choice))}">${escapeHTML(displayWord(choice))}</button>`).join("")}
      </div>
    </div>
  `;
}

function renderSightSentenceMatch(activity, index) {
  const word = activity?.word ?? "";
  const answer = meaningfulSightSentence(word);
  const distractors = sightSentenceDistractors(word);
  const choices = seededShuffle(uniqueWords([answer, ...distractors]).slice(0, 4), `${state.activeLearnerId}:${todayKey()}:sentence-match:${word}:${index}`);
  return `
    <div class="sight-game-panel">
      ${audioButton("Hear word", word, `sight-match-${normalizeWord(word)}`, "primary-action hear-word-action")}
      <p class="sight-helper">Which sentence uses <strong>${escapeHTML(displayWord(word))}</strong> in a good way?</p>
      <div class="sight-choice-grid sentence-choice-grid">
        ${choices.map((choice) => `<button class="sight-choice" type="button" data-sight-sentence-match="${escapeHTML(choice === answer ? "yes" : "no")}">${escapeHTML(choice)}</button>`).join("")}
      </div>
    </div>
  `;
}

function renderSightHunt(word) {
  const passageWords = sightHuntPassage(word, meaningfulSightSentence(word));
  return `
    <div class="sight-game-panel">
      ${audioButton("Hear word", word, `sight-hunt-${normalizeWord(word)}`, "primary-action hear-word-action")}
      <p class="sight-helper">Find <strong>${escapeHTML(displayWord(word))}</strong> in the sentence.</p>
      <p class="mini-passage hunt-passage">
        ${passageWords.map((item, index) => item.word
          ? `<button type="button" data-sight-hunt-word="${escapeHTML(normalizeWord(item.word))}" data-hunt-index="${index}">${escapeHTML(item.text)}</button>`
          : `<span>${escapeHTML(item.text)}</span>`
        ).join("")}
      </p>
    </div>
  `;
}

function sightHuntPassage(word, example) {
  const target = displayWord(word);
  const fallback = meaningfulSightSentence(target);
  const source = new RegExp(`\\b${target.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i").test(example)
    ? example
    : fallback;
  const parts = source.split(/(\b[A-Za-z']+\b)/g);
  return parts.map((text) => (/^[A-Za-z']+$/.test(text) ? { text, word: text } : { text, word: "" }));
}

function sentenceContainsWord(sentence, word) {
  const target = displayWord(word).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`\\b${target}\\b`, "i").test(sentence);
}

function meaningfulSightSentence(word, plan = activePlan()) {
  const target = displayWord(word);
  const key = normalizeWord(target);
  const storyTitle = plan.readingArticleTitle || plan.readingTitle || "the story";
  const topic = plan.uoiTopic || "class";
  const entryExample = weeklyVocabularyEntry(target)?.example || wordHelp[key]?.[1] || "";
  if (entryExample && sentenceContainsWord(entryExample, target)) return entryExample;
  const bank = {
    a: "A small idea can grow into a story.",
    after: `After lunch, we read ${storyTitle}.`,
    all: "All the children listened to the story.",
    an: "Joaquin saw an object and imagined an animal.",
    and: "Joaquin looked and imagined a zoo.",
    are: "The animals are part of his idea.",
    as: "He moved as quietly as a cat.",
    bear: "The bear walked slowly by the tree.",
    before: `Before reading, we look at the title.`,
    but: "The box was plain, but the children had an idea.",
    can: "Can you find the word in the sentence?",
    do: "What do you notice in the picture?",
    each: "Each child shared one idea.",
    for: "The class made a mural for the hallway.",
    from: "The idea came from a small object.",
    go: "We go back to the story for details.",
    had: "Joaquin had a new idea.",
    has: "The story has a beginning, middle, and end.",
    he: "He imagined animals at the zoo.",
    her: "She put her book on the table.",
    him: "The teacher asked him to explain his idea.",
    his: "Joaquin used his imagination.",
    boy: "The boy shared his idea with the class.",
    down: "The toy truck rolled down the ramp.",
    found: "Owen found a safe way to move the box.",
    how: "Tell how the wheel helps the cart move.",
    if: "If we listen carefully, we can answer.",
    in: "The answer is in the sentence.",
    is: "The story is about imagination.",
    it: "It looked like a real animal to Joaquin.",
    like: "The object looked like an animal.",
    look: "Look carefully at the sentence.",
    make: "The children make a plan together.",
    many: "Many ideas can help a story grow.",
    more: "Tell one more detail from the story.",
    my: "My favorite part is the ending.",
    not: "The object is not a real animal.",
    of: "The title of the story gives a clue.",
    on: "The book is on the desk.",
    one: "One detail helped me understand the story.",
    or: "Is the text fiction or informational?",
    now: "Now we can test the design.",
    out: "The class walked out to test the ramp.",
    said: "Joaquin said he could imagine a zoo.",
    see: "I see one detail in the picture.",
    she: "She read the sentence again.",
    so: "The room was quiet, so everyone could listen.",
    some: "Some words are tricky at first.",
    that: "That detail helped me understand.",
    the: "The story has a problem and a solution.",
    their: "The children shared their ideas.",
    them: "The teacher asked them to look again.",
    then: "Then, Joaquin imagined the object as an animal.",
    there: "There is a clue in the sentence.",
    these: "These words help us read smoothly.",
    they: "They listened to the story together.",
    this: "This sentence tells a complete idea.",
    time: "It was time to read.",
    to: "We listen to the story first.",
    two: "Two details support the answer.",
    up: "The child picked up the book.",
    use: "Use the word in a clear sentence.",
    was: "Joaquin was thinking about the zoo.",
    we: "We read the sentence together.",
    were: "The children were ready to answer.",
    what: "What happened first in the story?",
    when: "When did the character change?",
    which: "Which answer makes sense?",
    will: "We will read the next sentence.",
    with: "Read with a clear voice.",
    would: "Joaquin would imagine a new animal.",
    write: "Write one complete sentence.",
    you: "You can try the next word.",
    your: "Use your finger to track the words.",
    color: "The artist chose a bright color for the sky.",
    first: "First, Joaquin looked carefully at the object.",
    happy: "I felt happy when my picture was finished.",
    money: "The class counted money for the school fair.",
    music: "Soft music played while the children painted.",
    next: "Next, the class talked about the problem.",
    second: "The second idea helped the group finish.",
    sound: "The sound of rain tapped on the window.",
    then: "Then, the character tried a new idea.",
    without: "We can listen without shouting.",
    because: "Joaquin smiled because he had a new idea.",
    world: "Books help us learn about the world."
  };
  if (bank[key]) return bank[key];
  if (key.length <= 3) return `The class read the word ${target} while learning about ${topic}.`;
  return `The class talked about ${target} during our ${topic} lesson.`;
}

function meaningfulSightSentenceBlank(word) {
  const target = displayWord(word);
  return meaningfulSightSentence(target).replace(new RegExp(`\\b${target}\\b`, "i"), "____");
}

function sightSentenceDistractors(word) {
  return [
    "The class tested the ramp carefully.",
    "Mia gave a clear answer with one detail.",
    "Luis watched the wheel turn around the axle.",
    "Nora shared her opinion with a reason.",
    "The children celebrated helpful teamwork."
  ].filter((sentence) => !sentenceContainsWord(sentence, word)).slice(0, 3);
}

function rememberWrongWord(word) {
  const progress = activeProgress();
  const key = normalizeWord(word);
  if (!key) return;
  if (!progress.wrongWordProfile) progress.wrongWordProfile = {};
  const current = progress.wrongWordProfile[key] ?? {
    word: displayWord(word),
    missCount: 0,
    correctAfterMissCount: 0,
    status: "needs practice",
    sourceWeek: weekKey()
  };
  progress.wrongWordProfile[key] = {
    ...current,
    word: current.word || displayWord(word),
    missCount: (current.missCount ?? 0) + 1,
    correctAfterMissCount: 0,
    status: "needs practice",
    lastMissedAt: new Date().toISOString(),
    sourceWeek: current.sourceWeek || weekKey(),
    nextReviewAt: addDays(new Date(), 1)
  };
  if (!progress.weakSkills.includes("spelling")) progress.weakSkills.unshift("spelling");
}

function recordCorrectForWrongWord(word) {
  const progress = activeProgress();
  const key = normalizeWord(word);
  const current = progress.wrongWordProfile?.[key];
  if (!current) return;
  const correctAfterMissCount = (current.correctAfterMissCount ?? 0) + 1;
  progress.wrongWordProfile[key] = {
    ...current,
    correctAfterMissCount,
    lastCorrectAt: new Date().toISOString(),
    status: correctAfterMissCount >= 2 ? "improving" : current.status
  };
}

function randomSpellingSuccess() {
  const options = [
    { title: "Yes!", icon: "star", speakText: "Yes!" },
    { title: "Nice!", icon: "trophy", speakText: "Nice!" },
    { title: "Got it!", icon: "spark", speakText: "Got it!" },
    { title: "Correct!", icon: "target", speakText: "Correct!" },
    { title: "Great!", icon: "bright-star", speakText: "Great!" }
  ];
  return options[Math.floor(Math.random() * options.length)];
}

function showSpellingFeedback(type, title, message, { icon = "star", speakText = title, word = "" } = {}) {
  state.spellingFeedbackOverlay = { type, title, message, icon, speakText, word };
  if (speakText) {
    window.setTimeout(() => playAudio(speakText, `spelling-feedback-${type}-${Date.now()}`), 80);
  }
}

function clearSpellingFeedbackSoon(delay = 1100) {
  window.setTimeout(() => {
    state.spellingFeedbackOverlay = null;
    renderPracticeRoute();
  }, delay);
}

function advanceSpellingWord(delay = 1500) {
  window.setTimeout(() => {
    const session = ensureSpellingSession();
    const words = session.words ?? [];
    const nextIndex = (session.currentIndex ?? 0) + 1;
    const shouldComplete = !words.length || session.completed || spellingSessionCompletedCount(session) >= words.length || nextIndex >= words.length;
    if (shouldComplete) {
      session.completed = true;
      session.currentIndex = Math.max(words.length - 1, 0);
      session.attemptsForCurrent = 0;
      session.currentAnswer = "";
      session.updatedAt = new Date().toISOString();
      state.spellingFeedbackOverlay = null;
      state.spellingWordPulse = false;
      stopAudio();
      saveStore();
      renderAll();
      return;
    }
    session.currentIndex = nextIndex;
    session.attemptsForCurrent = 0;
    session.currentAnswer = "";
    session.updatedAt = new Date().toISOString();
    state.spellingFeedbackOverlay = null;
    state.spellingWordPulse = true;
    stopAudio();
    saveStore();
    renderAll();
    window.setTimeout(() => {
      state.spellingWordPulse = false;
      renderPracticeRoute();
    }, 650);
  }, delay);
}

function completeCurrentSpellingWord(word, correct, attemptsUsed) {
  const session = ensureSpellingSession();
  const key = normalizeWord(word);
  const alreadyCompleted = session.completedWords?.some((item) => normalizeWord(item) === key);
  if (!alreadyCompleted) {
    session.completedWords = [...(session.completedWords ?? []), word];
    recordPracticeAction("listen", { correct, items: 1, seconds: 20, render: false });
  }
  session.results = [
    ...(session.results ?? []),
    { word, correct, attempts: attemptsUsed, completedAt: new Date().toISOString() }
  ];
  session.currentAnswer = "";
  if (spellingSessionCompletedCount(session) >= (session.words?.length ?? 0)) {
    session.completed = true;
  }
  session.updatedAt = new Date().toISOString();
}

function handleSpellingInput(event) {
  const session = ensureSpellingSession();
  const value = String(event.target.value || "").replace(/[^A-Za-z']/g, "").toLowerCase();
  if (event.target.value !== value) event.target.value = value;
  session.currentAnswer = value;
  session.updatedAt = new Date().toISOString();
  saveStore();
}

function handleSpellingSubmit(event) {
  event?.preventDefault?.();
  const session = ensureSpellingSession();
  const words = session.words ?? [];
  const word = words[session.currentIndex ?? 0];
  if (!word) return;
  const inputAnswer = document.querySelector("#spellingAnswerInput")?.value;
  if (inputAnswer != null) session.currentAnswer = inputAnswer;
  const answer = normalizeWord(session.currentAnswer ?? "");
  if (!answer) return;
  const expected = normalizeWord(word);
  const correct = answer === expected;
  session.attemptsForCurrent = (session.attemptsForCurrent ?? 0) + 1;

  if (correct) {
    recordCorrectForWrongWord(word);
    completeCurrentSpellingWord(word, true, session.attemptsForCurrent);
    const success = randomSpellingSuccess();
    showSpellingFeedback("correct", success.title, "", { icon: success.icon, speakText: success.speakText, word });
    saveStore();
    renderPracticeRoute();
    advanceSpellingWord(1500);
    return;
  }

  rememberWrongWord(word);
  session.currentAnswer = "";
  if (session.attemptsForCurrent < MAX_SPELLING_ATTEMPTS) {
    showSpellingFeedback("try", "Try again.", "Listen one more time.", { icon: "retry", speakText: "Try again." });
    saveStore();
    renderPracticeRoute();
    window.setTimeout(() => speakSpellingPrompt(word, `retry-${expected}-${session.attemptsForCurrent}`), 950);
    clearSpellingFeedbackSoon(1300);
    return;
  }

  completeCurrentSpellingWord(word, false, session.attemptsForCurrent);
  showSpellingFeedback("saved", "Saved for practice.", "You will see this word again.", { icon: "notebook", speakText: "Saved for practice.", word });
  saveStore();
  renderPracticeRoute();
}

function recordSightWord(word, correct) {
  const progress = activeProgress();
  const key = normalizeWord(word);
  if (!progress.sightWordProfile) progress.sightWordProfile = {};
  const current = progress.sightWordProfile[key] ?? { word: displayWord(word), correct: 0, missed: 0, missCount: 0, lastSeen: "" };
  progress.sightWordProfile[key] = {
    ...current,
    word: current.word || displayWord(word),
    correct: (current.correct ?? 0) + (correct ? 1 : 0),
    missed: (current.missed ?? 0) + (correct ? 0 : 1),
    missCount: (current.missCount ?? current.missed ?? 0) + (correct ? 0 : 1),
    lastSeen: todayKey(),
    lastMissedAt: correct ? current.lastMissedAt : new Date().toISOString(),
    lastCorrectAt: correct ? new Date().toISOString() : current.lastCorrectAt,
    status: correct ? "building automaticity" : "needs practice"
  };
  if (!correct && !progress.weakSkills.includes("sight words")) progress.weakSkills.unshift("sight words");
}

function recordGrammarSkill(skill, correct) {
  const progress = activeProgress();
  const key = normalizeWord(skill || "grammar");
  if (!key) return;
  if (!progress.grammarSkillProfile) progress.grammarSkillProfile = {};
  const current = progress.grammarSkillProfile[key] ?? { skill: key, correct: 0, missed: 0, lastSeen: "" };
  progress.grammarSkillProfile[key] = {
    ...current,
    skill: current.skill || key,
    correct: (current.correct ?? 0) + (correct ? 1 : 0),
    missed: (current.missed ?? 0) + (correct ? 0 : 1),
    lastSeen: todayKey(),
    lastMissedAt: correct ? current.lastMissedAt : new Date().toISOString(),
    lastCorrectAt: correct ? new Date().toISOString() : current.lastCorrectAt
  };
  if (!correct && !progress.weakSkills.includes("grammar")) progress.weakSkills.unshift("grammar");
}

function advanceSightWord(delay = 900) {
  window.setTimeout(() => {
    const session = ensureSightSession();
    session.currentIndex = Math.min((session.currentIndex ?? 0) + 1, Math.max((session.activities?.length ?? 1) - 1, 0));
    session.buildState = null;
    session.updatedAt = new Date().toISOString();
    state.sightFeedbackOverlay = null;
    stopAudio();
    saveStore();
    renderAll();
  }, delay);
}

function completeSightWord(word, correct) {
  const session = ensureSightSession();
  const activity = session.activities?.[session.currentIndex ?? 0] ?? { word, type: "hear-find" };
  const activityId = activity.id ?? `${normalizeWord(word)}-${session.currentIndex ?? 0}`;
  const key = normalizeWord(word);
  recordSightWord(word, correct);
  if (correct && !session.completedActivities?.includes(activityId)) {
    session.completedActivities = [...(session.completedActivities ?? []), activityId];
    if (!session.completedWords?.some((item) => normalizeWord(item) === key)) {
      session.completedWords = [...(session.completedWords ?? []), word];
    }
    recordPracticeAction("sight", { correct: true, items: 1, seconds: 20, render: false });
  }
  session.results = [...(session.results ?? []), { word, type: activity.type, correct, completedAt: new Date().toISOString() }];
  session.updatedAt = new Date().toISOString();
  if (!correct) session.buildState = null;
  state.sightFeedbackOverlay = correct
    ? { type: "correct", title: "Nice reading!", message: "Next word is coming.", icon: "star" }
    : { type: "try", title: "Look again.", message: "Try the word one more time.", icon: "retry" };
  saveStore();
  renderPracticeRoute();
  if (correct) {
    playAudio("Nice reading!", `sight-good-${Date.now()}`);
    advanceSightWord(1100);
  } else {
    playAudio(displayWord(word), `sight-retry-${Date.now()}`);
    window.setTimeout(() => {
      state.sightFeedbackOverlay = null;
      renderPracticeRoute();
    }, 1200);
  }
}

function handleSightAction(target) {
  const { session, word, activity } = currentSightSessionState();
  if (!word) return;
  const expected = normalizeWord(word);
  const choice = target.closest("[data-sight-choice]");
  if (choice) {
    completeSightWord(word, choice.dataset.sightChoice === expected);
    return;
  }
  const sentenceMatch = target.closest("[data-sight-sentence-match]");
  if (sentenceMatch) {
    completeSightWord(word, sentenceMatch.dataset.sightSentenceMatch === "yes");
    return;
  }
  const fillChoice = target.closest("[data-sight-fill]");
  if (fillChoice) {
    const chunks = sightWordChunks(word);
    const answer = chunks[sightFocusPartIndex(chunks)] ?? "";
    completeSightWord(word, normalizeWord(fillChoice.dataset.sightFill) === normalizeWord(answer));
    return;
  }
  const buildLetter = target.closest("[data-sight-build-letter]");
  if (buildLetter) {
    const key = normalizeWord(word);
    const targetParts = sightWordChunks(word);
    const tiles = seededShuffle(targetParts, `${state.activeLearnerId}:${todayKey()}:build:${word}:${activity?.id ?? ""}`);
    if (!session.buildState || session.buildState.word !== key) {
      session.buildState = { word: key, parts: [], usedIndexes: [] };
    }
    const tileIndex = Number(buildLetter.dataset.sightBuildLetter);
    if (session.buildState.usedIndexes?.includes(tileIndex)) return;
    session.buildState.parts = [...(session.buildState.parts ?? session.buildState.letters ?? []), tiles[tileIndex]];
    session.buildState.usedIndexes = [...(session.buildState.usedIndexes ?? []), tileIndex];
    session.updatedAt = new Date().toISOString();
    const builtWord = normalizeWord(session.buildState.parts.join(""));
    saveStore();
    if (session.buildState.parts.length >= targetParts.length) {
      completeSightWord(word, builtWord === expected);
    } else {
      renderPracticeRoute();
    }
    return;
  }
  if (target.closest("[data-sight-build-clear]")) {
    session.buildState = { word: expected, parts: [], usedIndexes: [] };
    session.updatedAt = new Date().toISOString();
    saveStore();
    renderPracticeRoute();
    return;
  }
  const huntWord = target.closest("[data-sight-hunt-word]");
  if (huntWord) {
    completeSightWord(word, huntWord.dataset.sightHuntWord === expected);
    return;
  }
  if (target.closest("[data-sight-read]")) {
    completeSightWord(word, true);
  }
}

function handleReadingQuizChoice(target) {
  const choiceButton = target.closest("[data-reading-quiz-choice]");
  if (!choiceButton) return;
  const session = currentQuizSessionForActivePractice();
  if (!session) return;
  const story = session.story ?? currentListeningStory();
  if (session.completed || session.pendingFeedback) return;
  const questions = story.questions;
  const index = Math.min(session.currentIndex ?? 0, Math.max(questions.length - 1, 0));
  const question = questions[index];
  const selectedIndex = Number(choiceButton.dataset.readingQuizChoice);
  const correct = selectedIndex === question.answerIndex;
  session.pendingFeedback = {
    prompt: question.prompt,
    selectedIndex,
    answerIndex: question.answerIndex,
    selected: question.choices[selectedIndex],
    answer: question.answer,
    correct
  };
  session.updatedAt = new Date().toISOString();
  saveStore();
  renderAll();
}

function handleReadingQuizNext() {
  const session = currentQuizSessionForActivePractice();
  if (!session) return;
  const story = session.story ?? currentListeningStory();
  if (session.completed || !session.pendingFeedback) return;
  const questions = story.questions;
  const feedback = session.pendingFeedback;
  session.answers = [
    ...(session.answers ?? []),
    {
      prompt: feedback.prompt,
      selected: feedback.selected,
      answer: feedback.answer,
      correct: feedback.correct
    }
  ];
  session.score = (session.score ?? 0) + (feedback.correct ? 1 : 0);
  session.currentIndex = (session.currentIndex ?? 0) + 1;
  session.pendingFeedback = null;
  session.updatedAt = new Date().toISOString();
  if (session.currentIndex >= questions.length) {
    session.completed = true;
    if (state.activePracticeId === "story-reward") {
      if (!activeProgress().storyRewardQuizAttempts) activeProgress().storyRewardQuizAttempts = [];
      activeProgress().storyRewardQuizAttempts.push({
        date: todayKey(),
        storyTitle: story.storyTitle,
        score: session.score,
        total: questions.length,
        answers: session.answers,
        completedAt: new Date().toISOString()
      });
    } else {
      if (!activeProgress().readingQuizAttempts) activeProgress().readingQuizAttempts = [];
      activeProgress().readingQuizAttempts.push({
        date: todayKey(),
        storyTitle: story.storyTitle,
        score: session.score,
        total: questions.length,
        answers: session.answers,
        completedAt: new Date().toISOString()
      });
      recordPracticeAction("reading", { correct: session.score >= Math.ceil(questions.length * 0.6), items: 1, seconds: 90, render: false });
    }
  }
  saveStore();
  renderAll();
}

function handleGrammarChoice(target) {
  const choiceButton = target.closest("[data-grammar-choice]");
  if (!choiceButton) return;
  const session = ensureGrammarSession();
  if (session.completed || session.pendingFeedback) return;
  const questions = session.questions ?? grammarQuestions();
  const index = Math.min(session.currentIndex ?? 0, Math.max(questions.length - 1, 0));
  const question = questions[index];
  const selectedIndex = Number(choiceButton.dataset.grammarChoice);
  const correct = selectedIndex === question.answerIndex;
  recordGrammarSkill(question.skill || "grammar", correct);
  session.pendingFeedback = {
    selectedIndex,
    answerIndex: question.answerIndex,
    correct
  };
  session.answers = [
    ...(session.answers ?? []),
    {
      prompt: question.prompt,
      selected: question.choices[selectedIndex],
      answer: question.answer,
      correct
    }
  ];
  session.score = (session.score ?? 0) + (correct ? 1 : 0);
  session.updatedAt = new Date().toISOString();
  saveStore();
  renderAll();
  window.setTimeout(() => {
    if (!session.pendingFeedback) return;
    session.pendingFeedback = null;
    session.currentIndex = index + 1;
    session.updatedAt = new Date().toISOString();
    if (session.currentIndex >= questions.length) {
      session.completed = true;
      if (!activeProgress().grammarPracticeAttempts) activeProgress().grammarPracticeAttempts = [];
      activeProgress().grammarPracticeAttempts.push({
        date: todayKey(),
        grammarFocus: activePlan().grammarFocus || "",
        score: session.score,
        total: questions.length,
        answers: session.answers,
        completedAt: new Date().toISOString()
      });
      recordPracticeAction("wordwork", { correct: session.score >= Math.ceil(questions.length * 0.6), items: questions.length, seconds: 90, render: false });
    }
    saveStore();
    renderAll();
  }, 1100);
}

function renderWeeklyForm() {
  const plan = activePlan();
  const editing = Boolean(state.weeklyPlanEditing);
  const weekInfo = planWeekInfo(plan);
  const fields = [
    ["#readingTitleInput", plan.readingTitle],
    ["#phonicsFocusInput", plan.phonicsFocus],
    ["#grammarFocusInput", plan.grammarFocus],
    ["#uoiTopicInput", plan.uoiTopic],
    ["#masteryWordsInput", plan.masteryWords.join(", ")],
    ["#sightWordsInput", plan.sightWords.join(", ")],
    ["#epicIdeasInput", plan.epicIdeas.join(", ")]
  ];

  document.querySelector("#currentWeekLabel").textContent = `Current Week: ${weekInfo.calendarWeekLabel || weekInfo.weekId}`;
  fields.forEach(([selector, value]) => {
    const field = document.querySelector(selector);
    if (!field) return;
    field.value = value ?? "";
    field.disabled = !editing;
  });

  document.querySelector("#weeklyForm")?.classList.toggle("weekly-plan-readonly", !editing);
  const editButton = document.querySelector("#editWeeklyButton");
  const saveButton = document.querySelector("#saveWeeklyButton");
  const resetButton = document.querySelector("#resetWeeklyButton");
  const status = document.querySelector("#weeklyEditStatus");
  if (editButton) editButton.disabled = editing;
  if (saveButton) saveButton.disabled = !editing;
  if (resetButton) resetButton.disabled = !editing;
  if (status) status.textContent = editing ? "Editing current week plan." : "View only. Select Edit to make changes.";
}

function renderFridayCheck() {
  document.querySelector("#fridayWordGrid").innerHTML = activePlan().masteryWords
    .map((word) => {
      const key = normalizeWord(word);
      const stat = activeProgress().wordStats[key] ?? { correct: 0, attempts: 0 };
      const definition = wordHelp[key]?.[0] ?? "a weekly word to read, spell, and use";
      return `
        <article class="word-card">
          <div>
            <button class="word-title word-link" type="button" data-word="${key}">${word}</button>
            <p>${clickableText(definition)}</p>
          </div>
          <span class="score-pill">${stat.correct}/${stat.attempts} correct</span>
          <div class="card-actions">
            ${audioButton("Hear word", word, `friday-${key}`)}
            <button class="primary-action" type="button" data-word-result="${word}:correct">Correct</button>
            <button class="secondary-action" type="button" data-word-result="${word}:practice">Practice again</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderReview() {
  const due = dueReviewWords();
  const upcoming = Object.entries(activeProgress().reviewQueue)
    .filter(([word]) => !due.some((item) => item.word === word))
    .slice(0, 8);

  document.querySelector("#reviewList").innerHTML = `
    <section class="review-block">
      <h3>Due today</h3>
      ${
        due.length
          ? due
              .map(
                (item) => `
                  <article class="review-row">
                    <button class="word-link" type="button" data-word="${item.word}">${item.word}</button>
                    <span>Stage ${item.stage + 1}</span>
                    <button class="primary-action" type="button" data-word-result="${item.word}:correct">Reviewed</button>
                  </article>
                `
              )
              .join("")
          : "<p>No review words are due today.</p>"
      }
    </section>
    <section class="review-block">
      <h3>Coming next</h3>
      ${
        upcoming.length
          ? upcoming
              .map(
                ([word, item]) => `
                  <article class="review-row">
                    <button class="word-link" type="button" data-word="${word}">${word}</button>
                    <span>${item.nextReview}</span>
                  </article>
                `
              )
              .join("")
          : "<p>Practice Friday words to start the review schedule.</p>"
      }
    </section>
  `;
}

function renderSkills() {
  document.querySelector("#skillColumns").innerHTML = skillPath
    .map(
      (band) => `
        <article class="skill-column">
          <p class="eyebrow">${band.subtitle}</p>
          <h3>${band.title}</h3>
          <ul>
            ${band.skills.map((skill) => `<li>${clickableText(skill)}</li>`).join("")}
          </ul>
        </article>
      `
    )
    .join("");
}

function renderPractice() {
  const select = document.querySelector("#practiceDomainSelect");
  select.innerHTML = practiceDomains.map((domain) => `<option value="${domain.id}">${domain.title}</option>`).join("");
  select.value = state.practiceDomain;

  const domain = practiceDomains.find((item) => item.id === state.practiceDomain) ?? practiceDomains[0];
  document.querySelector("#practiceBoard").innerHTML = domain.items
    .map(([title, prompt], index) => {
      const spellingWord = activePlan().masteryWords[index % Math.max(activePlan().masteryWords.length, 1)] ?? "";
      const targetAudio = title.toLowerCase().includes("listen") && spellingWord
        ? audioButton("Hear word", spellingWord, `practice-${index}-${normalizeWord(spellingWord)}`)
        : "";
      return `
        <article class="practice-card">
          <h3>${title}</h3>
          <p>${clickableText(prompt)}</p>
          <div class="card-actions">
            ${targetAudio}
            <button class="primary-action" type="button" data-skill-practice="${title}">Practice done</button>
          </div>
        </article>
      `;
    })
    .join("");
}

function renderExam() {
  document.querySelector("#examGrid").innerHTML = examItems
    .map(
      (item, index) => `
        <article class="exam-card">
          <p class="eyebrow">${item.title}</p>
          <h3>${clickableText(item.prompt)}</h3>
          <div class="choice-list">
            ${item.choices
              .map(
                (choice) => `
                  <label class="choice">
                    <input type="radio" name="exam-${index}" value="${choice}">
                    <span>${clickableText(choice)}</span>
                  </label>
                `
              )
              .join("")}
          </div>
          <p class="answer-note" id="exam-note-${index}"></p>
        </article>
      `
    )
    .join("");
}

function renderCloudSyncPanel() {
  const signedIn = cloudSyncAvailable() && window.ReadingBridgeCloud.isSignedIn();
  const familyDocId = signedIn ? window.ReadingBridgeCloud.familyDocId?.() : "";
  return `
    <article class="parent-card cloud-sync-card">
      <div class="parent-card-head">
        <span class="profile-avatar small" aria-hidden="true">☁</span>
        <div>
          <p class="eyebrow">Cloud sync</p>
          <h3>Hearts and progress</h3>
        </div>
      </div>
      <div class="cloud-sync-status-row">
        <span class="status-pill" id="cloudSyncStatus">${escapeHTML(state.cloudSyncStatus)}</span>
        ${familyDocId ? `<span>${escapeHTML(familyDocId)}</span>` : ""}
      </div>
      <p class="material-status cloud-sync-message" id="cloudSyncMessage">${escapeHTML(state.cloudSyncMessage)}</p>
      <p>This device connects automatically. Peggy and Gina sync across devices; Tester stays local.</p>
      <div class="card-actions">
        <button class="secondary-action" type="button" data-cloud-sync-now ${signedIn ? "" : "disabled"}>Sync now</button>
      </div>
    </article>
  `;
}

function renderParentSettingsPanel() {
  return `
    <details class="parent-settings-menu">
      <summary class="parent-settings-toggle" aria-label="Parent settings" title="Parent settings">
        <span aria-hidden="true">⚙</span>
      </summary>
      <div class="parent-settings-panel">
        <article class="parent-card parent-session-card">
          <div>
            <p class="eyebrow">Secure parent session</p>
            <h3>Parent is unlocked</h3>
            <p>Local materials and parent setup are available for this browser session.</p>
          </div>
          <div class="card-actions">
            <button class="secondary-action" id="parentLogoutButton" type="button">Log out</button>
          </div>
        </article>
        ${renderCloudSyncPanel()}
        <article class="parent-card password-card">
          <p class="eyebrow">Password</p>
          <h3>Change parent password</h3>
          <form class="password-form" id="parentPasswordForm">
            <label>
              Current password
              <input id="currentParentPasswordInput" type="password" autocomplete="current-password" />
            </label>
            <label>
              New password
              <input id="newParentPasswordInput" type="password" autocomplete="new-password" />
            </label>
            <button class="primary-action" type="submit">Change password</button>
          </form>
          <p class="material-status" id="passwordStatus"></p>
        </article>
      </div>
    </details>
  `;
}

function renderParent() {
  const learnerCards = learners
    .map((learner) => {
      const progress = store.learners[learner.id];
      const completedToday = progress.centerCompletions[todayKey()]?.length ?? 0;
      const weak = progress.weakSkills.slice(0, 5);
      const mastered = Object.values(progress.wordStats).filter((stat) => stat.mastered).length;
      const due = Object.values(progress.reviewQueue).filter((item) => item.nextReview <= todayKey()).length;
      const recentEpic = progress.epicLog.slice(-3).reverse();
      const wrongWords = Object.values(progress.wrongWordProfile ?? {})
        .sort((a, b) => (b.missCount ?? 0) - (a.missCount ?? 0))
        .slice(0, 6);
      const sightWrongWords = Object.values(progress.sightWordProfile ?? {})
        .filter((item) => (item.missed ?? 0) > 0)
        .sort((a, b) => (b.missCount ?? b.missed ?? 0) - (a.missCount ?? a.missed ?? 0))
        .slice(0, 6);
      return `
        <article class="parent-card learner-summary-card">
          <div class="parent-card-head">
            <span class="profile-avatar small" aria-hidden="true">${escapeHTML(profileName(learner.id).slice(0, 1).toUpperCase())}</span>
            <div>
              <p class="eyebrow">Learner summary</p>
              <h3>${escapeHTML(profileName(learner.id))}</h3>
            </div>
          </div>
          <div class="summary-strip">
            <span><strong>${completedToday}</strong> centers today</span>
            <span><strong>${mastered}</strong> words mastered</span>
            <span><strong>${due}</strong> reviews due</span>
          </div>
          <h4>Weakness focus</h4>
          <ul>${weak.map((skill) => `<li>${clickableText(skill)}</li>`).join("")}</ul>
          <h4>Spelling watchlist</h4>
          ${
            wrongWords.length
              ? `<ul>${wrongWords.map((item) => `<li>${escapeHTML(item.word)} · missed ${item.missCount ?? 1}</li>`).join("")}</ul>`
              : "<p>No spelling misses yet.</p>"
          }
          <h4>Sight word watchlist</h4>
          ${
            sightWrongWords.length
              ? `<ul>${sightWrongWords.map((item) => `<li>${escapeHTML(item.word)} · missed ${item.missCount ?? item.missed ?? 1}</li>`).join("")}</ul>`
              : "<p>No sight word misses yet.</p>"
          }
          <h4>Epic reading</h4>
          ${
            recentEpic.length
              ? `<ul>${recentEpic.map((entry) => `<li>${entry.date}: ${clickableText(entry.book)} - ${clickableText(entry.note || "completed")}</li>`).join("")}</ul>`
              : "<p>No Epic reading logged yet.</p>"
          }
          <button class="secondary-action" type="button" data-open-learner="${learner.id}">View child practice</button>
        </article>
      `;
    })
    .join("");
  const shelf = store.parentBooks ?? [];
  const archives = store.weeklyArchives ?? [];
  const wordBankEntries = Object.values(store.wordTestBank ?? {}).sort((a, b) => (b.lastSeenAt || "").localeCompare(a.lastSeenAt || ""));
  const pendingPlan = store.pendingWeeklyPlan;
  document.querySelector("#parentGrid").innerHTML = `
    ${renderParentSettingsPanel()}
    ${learnerCards}
    <article class="parent-card weekly-archive-card">
      <p class="eyebrow">Weekly archive</p>
      <h3>${escapeHTML(planWeekInfo(activePlan()).archiveLabel)}</h3>
      <p><strong>Active plan:</strong> ${escapeHTML(activePlan().readingArticleTitle || activePlan().readingTitle || activePlan().weekId || "Weekly plan")}</p>
      <p><strong>Pending next week:</strong> ${pendingPlan ? escapeHTML(`${pendingPlan.calendarWeekLabel || pendingPlan.weekId}: ${pendingPlan.readingArticleTitle || pendingPlan.readingTitle || pendingPlan.sourceTitle || "Weekly plan"}`) : "None"}</p>
      <p>Save this week's plan, identified words, skills, and learner snapshots for parent meetings and future review.</p>
      <div class="card-actions">
        <button class="primary-action" id="archiveThisWeekButton" type="button">Archive this week</button>
        <button class="secondary-action" id="clearAllTestProgressButton" type="button">Clear all test practice records</button>
      </div>
      <div class="weekly-data-summary">
        <p><strong>Archives:</strong> ${archives.length}</p>
        <p><strong>Long-term word bank:</strong> ${wordBankEntries.length} words</p>
      </div>
    </article>
    <article class="parent-card archive-list-card">
      <p class="eyebrow">Parent meeting records</p>
      <h3>Weekly archives</h3>
      ${
        archives.length
          ? archives
              .slice(0, 8)
              .map(
                (archive) => `
                  <div class="archive-row">
                    <strong>${escapeHTML(archive.archiveLabel)}</strong>
                    <span>${escapeHTML(archive.weeklyPlanSnapshot?.readingArticleTitle || archive.weeklyPlanSnapshot?.readingTitle || "Weekly plan")}</span>
                    <small>${escapeHTML([
                      `${archive.identifiedWords?.highFrequencyWords?.length ?? 0} sight`,
                      `${archive.identifiedWords?.uoiVocabulary?.length ?? 0} UOI`,
                      `${archive.identifiedWords?.phonicsReviewWords?.length ?? 0} phonics`,
                      `${archive.identifiedWords?.grammarPracticeWords?.length ?? 0} grammar`
                    ].join(" · "))}</small>
                  </div>
                `
              )
              .join("")
          : "<p>No weekly archives yet.</p>"
      }
    </article>
    <article class="parent-card word-bank-card">
      <p class="eyebrow">Long-term word bank</p>
      <h3>Retained weekly words</h3>
      ${
        wordBankEntries.length
          ? `<div class="tag-cloud">${wordBankEntries.slice(0, 60).map((entry) => `<span title="${escapeHTML(entry.sourceArchiveLabel || "")}">${escapeHTML(entry.word)}</span>`).join("")}</div>`
          : "<p>No archived words yet.</p>"
      }
    </article>
    <article class="parent-card epic-shelf-card">
      <div class="parent-card-head">
        <span class="profile-avatar small" aria-hidden="true">E</span>
        <div>
          <p class="eyebrow">Epic shelf</p>
          <h3>Parent-curated books</h3>
        </div>
      </div>
      <form class="shelf-form" id="epicShelfForm">
        <label>
          Book title
          <input id="shelfTitleInput" type="text" placeholder="Type an Epic book title" />
        </label>
        <label>
          Topic / skill tags
          <input id="shelfTagsInput" type="text" placeholder="simple machines, long vowels, nonfiction" />
        </label>
        <button class="primary-action" type="submit">Add to shelf</button>
      </form>
      <div class="shelf-list">
        ${
          shelf.length
            ? shelf
                .map(
                  (book) => `
                    <article>
                      <strong>${escapeHTML(book.title)}</strong>
                      <span>${escapeHTML([...(book.topicTags ?? []), ...(book.skillTags ?? [])].join(", ") || "parent pick")}</span>
                    </article>
                  `
                )
                .join("")
            : "<p>No parent books saved yet. Add good Epic finds here after searching on the iPad.</p>"
        }
      </div>
    </article>
    <article class="parent-card library-card">
      <p class="eyebrow">Seed library</p>
      <h3>${epicSeedLibrary.length + epicSearchLibrary.length} starter entries</h3>
      <p>The matcher uses public metadata and curated G1/G2 tags. Add iPad finds to make it more precise for Peggy and Gina.</p>
      <div class="tag-cloud">
        <span>simple machines</span><span>architecture</span><span>long vowels</span><span>contractions</span><span>syllables</span><span>main idea</span>
      </div>
    </article>
  `;
}

function renderMaterials() {
  const container = document.querySelector("#materialsGrid");
  if (!container) return;
  if (!parentIsAuthenticated()) {
    container.innerHTML = `
      <section class="wide-panel setup-panel">
        <p class="eyebrow">Locked</p>
        <h3>Parent login required</h3>
        <p>Unlock Parent before viewing or uploading learning materials.</p>
        <button class="primary-action" type="button" data-open-parent-auth>Unlock Parent</button>
      </section>
    `;
    return;
  }

  const suggestion = state.weeklySuggestions;
  const currentWeekLabel = learningWeekInfo().calendarWeekLabel;
  const nextWeekLabel = nextWeekInfoForDate().calendarWeekLabel;
  const batch = state.materialBatch;
  const editableReviewField = (label, field, value, rows = 2) => `
    <label class="weekly-review-field">
      <span>${escapeHTML(label)}</span>
      <textarea data-weekly-review-field="${escapeHTML(field)}" rows="${rows}">${escapeHTML(value || "")}</textarea>
    </label>
  `;
  const suggestionReviewFields = suggestion ? [
    editableReviewField("UOI / Topic Vocabulary", "uiVocabulary", (suggestion.uiVocabulary ?? []).map((item) => item.word).join(", ")),
    editableReviewField("English Vocabulary", "englishVocabulary", (suggestion.englishVocabulary ?? []).map((item) => item.word).join(", "), 3),
    editableReviewField("Spelling Basic", "spellingBasic", (suggestion.spelling?.basic ?? []).join(", "), 3),
    editableReviewField("Spelling Review", "spellingReview", (suggestion.spelling?.review ?? []).join(", ")),
    editableReviewField("Spelling Challenge", "spellingChallenge", (suggestion.spelling?.challenge ?? []).join(", ")),
    editableReviewField("High frequency", "highFrequencyWords", (suggestion.highFrequencyWords ?? []).join(", ")),
    editableReviewField("Grammar practice", "grammarPracticeWords", (suggestion.grammarPracticeWords ?? []).join(", ")),
    editableReviewField("Skills", "skills", [suggestion.phonicsFocus, suggestion.grammarFocus, suggestion.readingFocus].filter(Boolean).join(", "), 3),
    editableReviewField("Reading texts", "readingSelections", (suggestion.readingSelections ?? []).join(", ") || suggestion.readingTitle || ""),
    editableReviewField("HMH source / search link", "readingArticleSourceUrl", suggestion.readingArticleSourceUrl || hmhArticleSearchUrl(suggestion.readingTitle || "")),
    editableReviewField("Poem details for Level 2", "poemSelections", formatPoemSelectionsForReview(suggestion.poemSelections ?? []), 5),
    editableReviewField("Main story details for retell", "readingArticleEvidence", suggestion.readingArticleEvidence || "", 5),
    editableReviewField("Story characters", "readingArticleCharacters", (suggestion.readingArticleCharacters ?? []).join(", ")),
    editableReviewField("Story setting", "readingArticleSetting", suggestion.readingArticleSetting || ""),
    editableReviewField("Story problem / challenge", "readingArticleProblem", suggestion.readingArticleProblem || "", 3),
    editableReviewField("Story key events", "readingArticleKeyEvents", (suggestion.readingArticleKeyEvents ?? []).join("\n"), 4),
    editableReviewField("Story ending", "readingArticleEnding", suggestion.readingArticleEnding || "", 3),
    editableReviewField("Story big idea", "readingArticleBigIdea", suggestion.readingArticleBigIdea || "", 3),
    editableReviewField("Learning points", "weeklyLearningOutcomes", [...(suggestion.weeklyLearningOutcomes ?? []), ...(suggestion.homeLearningSuggestions ?? [])].join("\n"), 4)
  ].join("") : "";

  container.innerHTML = `
    ${
      suggestion
        ? `<section class="wide-panel suggestion-panel">
            <p class="eyebrow">Review weekly plan</p>
            <h3>${escapeHTML(suggestion.sourceTitle)}</h3>
            <p>Choose whether this plan updates ${escapeHTML(currentWeekLabel)} now or waits for ${escapeHTML(nextWeekLabel)} on Sunday.</p>
            <p>Edit anything OCR got wrong before applying.</p>
            <div class="weekly-review-editor">
              ${suggestionReviewFields}
            </div>
            <div class="card-actions">
              <button class="primary-action" type="button" data-apply-material-suggestions="this-week">Use for this week</button>
              <button class="primary-action" type="button" data-apply-material-suggestions="next-week">Use for next week</button>
              <button class="secondary-action" type="button" id="dismissMaterialSuggestionsButton">Dismiss</button>
            </div>
          </section>`
        : ""
    }
    ${
      batch && !suggestion
        ? `<section class="wide-panel suggestion-panel">
            <p class="eyebrow">Uploaded batch</p>
            <h3>${escapeHTML(batch.title)}</h3>
            <p>${escapeHTML(`${batch.fileCount} file${batch.fileCount > 1 ? "s" : ""} uploaded. ${batch.parsedCount} parsed. ${batch.weeklySuggestionCount} weekly plan page${batch.weeklySuggestionCount === 1 ? "" : "s"} detected.`)}</p>
            <div class="card-actions">
              <button class="primary-action" type="button" data-create-weekly-plan>Create weekly plan</button>
              <button class="secondary-action" type="button" data-dismiss-material-batch>Dismiss</button>
            </div>
          </section>`
        : ""
    }
    <section class="wide-panel material-upload-panel">
      <div>
        <p class="eyebrow">Upload</p>
        <h3>Create weekly plan</h3>
        <p>Upload all weekly plan pages together. For Level 2 retell, upload HMH story pages with Material kind set to Main Story Pages.</p>
      </div>
      <form class="material-form" id="materialUploadForm">
        <label class="span-2">
          File
          <input id="materialFileInput" type="file" accept=".jpg,.jpeg,.png,.webp,.pdf,.xlsx,.xls,.csv" multiple />
        </label>
        <label>
          Title
          <input id="materialTitleInput" type="text" placeholder="S2 final review packet" />
        </label>
        <label>
          Source
          <input id="materialSourceInput" type="text" placeholder="school / 上中国际 / HMH" />
        </label>
        <label>
          Grade
          <select id="materialGradeInput">
            <option value="G1">G1</option>
            <option value="G2">G2</option>
            <option value="G3">G3</option>
            <option value="mixed">Mixed</option>
          </select>
        </label>
        <label>
          Semester
          <select id="materialSemesterInput">
            <option value="S1">S1</option>
            <option value="S2">S2</option>
            <option value="summer">Summer</option>
            <option value="unknown">Unknown</option>
          </select>
        </label>
        <label>
          Material kind
          <select id="materialKindInput">
            <option value="weekly_plan">Weekly plan</option>
            <option value="main_story_pages">Main Story Pages</option>
            <option value="exam_paper">Exam paper</option>
            <option value="review_packet">Review packet</option>
            <option value="worksheet">Worksheet</option>
            <option value="reading_list">Reading list</option>
            <option value="parent_note">Parent note</option>
          </select>
        </label>
        <label>
          Skill tags
          <input id="materialSkillTagsInput" type="text" placeholder="long vowels, syllables, contractions" />
        </label>
        <label class="span-2">
          Topic tags
          <input id="materialTopicTagsInput" type="text" placeholder="simple machines, architecture, reading comprehension" />
        </label>
        <label class="span-2">
          Notes
          <textarea id="materialNotesInput" placeholder="What should this material help us practice?"></textarea>
        </label>
        <button class="primary-action span-2" type="submit">Upload material</button>
      </form>
      <div class="card-actions">
        <button class="secondary-action" type="button" data-clear-uploaded-materials>Clear current weekly upload</button>
      </div>
      <p class="material-status">${escapeHTML(state.materialsStatus)}</p>
    </section>
  `;
}

function materialWeekTagSummary(material) {
  const text = [
    material.title,
    material.fileName,
    material.source,
    material.notes,
    ...(material.skillTags ?? []),
    ...(material.topicTags ?? []),
    ...(material.detectedSkills ?? [])
  ].join(" ");
  const matches = text.match(/\b(?:calendar\s*week|cw|week)\s*#?\s*\d{1,2}\b/gi) ?? [];
  const explicitIds = text.match(/\b\d{4}-cw\d{2}\b/gi) ?? [];
  const tags = uniqueWords([...matches, ...explicitIds]);
  return tags.length ? `Week tags: ${tags.join(", ")}` : "No structured calendar week tag";
}

function materialDisplayDetectedWords(material) {
  const weekly = cleanedWeeklyPlanData(material.weeklyPlanData ?? {});
  const structured = cleanOcrVocabularyList([
    ...(weekly.vocabulary ?? []),
    ...(weekly.uoiVocabulary ?? []),
    ...(weekly.uiVocabulary ?? []),
    ...(weekly.englishVocabulary ?? []),
    ...(weekly.phonicsVocabulary ?? []),
    ...(weekly.highFrequencyWords ?? []),
    ...(weekly.spellingBasicWords ?? []),
    ...(weekly.spellingReviewWords ?? []),
    ...(weekly.spellingChallengeWords ?? [])
  ]);
  return structured.length ? structured : cleanOcrVocabularyList(material.detectedWords ?? []);
}

function renderWeeklyPlanDataSummary(data) {
  if (!data) return "";
  const cleaned = cleanedWeeklyPlanData(data);
  const rows = [
    ["Reading Title", cleaned.readingTitle ? [cleaned.readingTitle] : []],
    ["Reading Texts", cleaned.readingSelections],
    ["Jump Rope Reader", cleaned.jumpRopeReader ? [cleaned.jumpRopeReader] : []],
    ["Vocabulary", cleaned.vocabulary],
    ["UOI Vocabulary", cleaned.uoiVocabulary],
    ["English Vocabulary", cleaned.englishVocabulary],
    ["Phonics Words", cleaned.phonicsVocabulary],
    ["High Frequency", cleaned.highFrequencyWords],
    ["Spelling", [...(cleaned.spellingBasicWords ?? []), ...(cleaned.spellingReviewWords ?? []), ...(cleaned.spellingChallengeWords ?? [])]],
    ["Grammar Practice", cleaned.grammarPracticeWords],
    ["Grammar", cleaned.grammar],
    ["Foundational", cleaned.foundationalSkills],
    ["Learning Outcomes", cleaned.weeklyLearningOutcomes],
    ["Home Learning", cleaned.homeLearningSuggestions]
  ].filter(([, values]) => values?.length);
  if (!rows.length) return "";
  return `<div class="weekly-data-summary">${rows.map(([label, values]) => `<p><strong>${label}:</strong> ${escapeHTML(values.join(", "))}</p>`).join("")}</div>`;
}

function renderSections() {
  const route = currentPracticeRoute();
  if (!visibleSections().includes(state.activeSection)) {
    state.activeSection = visibleSections()[0];
  }
  document.querySelectorAll(".section-view").forEach((section) => {
    section.classList.toggle("active", !route && section.id === `section-${state.activeSection}`);
  });
  document.querySelector("#practiceRouteView").hidden = !route;
  document.querySelectorAll(".nav-item").forEach((button) => {
    const visible = visibleSections().includes(button.dataset.section);
    button.hidden = !visible;
    button.classList.toggle("active", !route && button.dataset.section === state.activeSection);
  });
  const titles = {
    today: "Today's Literacy Mission",
    weekly: "Weekly Planner",
    friday: "Friday Word Check",
    review: "Spaced Review",
    skills: "Skill Path",
    practice: "Practice Centers",
    exam: "Exam Practice",
    materials: "Learning Materials",
    parent: "Parent Review"
  };
  document.querySelector("#pageTitle").textContent = route ? "" : titles[state.activeSection];
  document.querySelector("#roleEyebrow").textContent =
    state.activeRole === "parent" ? "Parent dashboard" : `${activeLearner().name}'s practice`;
  document.querySelector("#heroTitle").textContent =
    state.activeRole === "parent" ? "Progress, weekly setup, and next focus" : `Hi ${activeLearner().name}, ready for today?`;
  document.querySelector("#heroText").textContent =
    state.activeRole === "parent"
      ? "Review daily completion, weak skills, Friday words, spaced review, and Epic reading for both children."
      : "Pick up your stars one center at a time. Listen, read, speak, write, and finish with Epic reading.";
  const currentDate = document.querySelector("#currentDate");
  if (currentDate) {
    currentDate.querySelector(".date-value").textContent = displayDate();
    currentDate.classList.toggle("home-date", state.activeRole === "child" && state.activeSection === "today" && !route);
  }
  document.querySelector("#childSwitchButton").hidden = state.activeRole === "parent";
  document.querySelector("#clearProgressButton").hidden = state.activeRole === "parent";
}

function renderAll() {
  renderProfileGate();
  renderLearners();
  renderSections();
  renderSummary();
  renderToday();
  renderWeeklyForm();
  renderFridayCheck();
  renderReview();
  renderSkills();
  renderPractice();
  renderExam();
  renderParent();
  renderMaterials();
  renderPracticeRoute();
  renderHeartRewardOverlay();
}

function renderHeartRewardOverlay() {
  document.querySelector("#heartRewardOverlay")?.remove();
  if (!state.heartRewardOverlay) return;
  const overlay = document.createElement("div");
  overlay.className = "heart-reward-overlay";
  overlay.id = "heartRewardOverlay";
  overlay.setAttribute("role", "dialog");
  overlay.setAttribute("aria-modal", "true");
  overlay.setAttribute("aria-label", "Heart reward");
  overlay.innerHTML = `
    <div class="heart-reward-card">
      <div class="heart-burst" aria-hidden="true">♥</div>
      <p class="eyebrow">Daily practice complete</p>
      <h2>You earned a heart!</h2>
      <p>${heartCount()} hearts collected.</p>
      <button class="primary-action" type="button" data-close-heart-reward>OK</button>
    </div>
  `;
  document.body.appendChild(overlay);
}

function saveWeeklyPlan(event) {
  event.preventDefault();
  if (!state.weeklyPlanEditing) {
    renderWeeklyForm();
    return;
  }
  const weekInfo = learningWeekInfo();
  store.weeklyPlan = applyWeekMetadata({
    ...activePlan(),
    readingTitle: document.querySelector("#readingTitleInput").value.trim(),
    phonicsFocus: document.querySelector("#phonicsFocusInput").value.trim(),
    grammarFocus: document.querySelector("#grammarFocusInput").value.trim(),
    uoiTopic: document.querySelector("#uoiTopicInput").value.trim(),
    masteryWords: uniqueWords(document.querySelector("#masteryWordsInput").value.split(/,|\n/)),
    sightWords: uniqueWords(document.querySelector("#sightWordsInput").value.split(/,|\n/)),
    epicIdeas: uniqueWords(document.querySelector("#epicIdeasInput").value.split(/,|\n/)),
    updatedAt: new Date().toISOString()
  }, weekInfo, "active");
  store.currentWeekId = store.weeklyPlan.weekId;
  if (!store.weeklyPlansByWeekId) store.weeklyPlansByWeekId = {};
  store.weeklyPlansByWeekId[weekInfo.weekId] = store.weeklyPlan;
  state.weeklyPlanEditing = false;
  saveStore();
  renderAll();
}

function markCenterComplete(centerId) {
  recordPracticeAction(centerId);
}

function completeMission() {
  // Child completion is behavior-based; this remains a no-op for legacy bindings.
  renderAll();
}

function logEpic(event) {
  event.preventDefault();
  const book = document.querySelector("#epicBookInput").value.trim();
  const note = document.querySelector("#epicNoteInput").value.trim();
  if (!book) return;
  activeProgress().epicLog.push({ date: todayKey(), book, note });
  document.querySelector("#epicBookInput").value = "";
  document.querySelector("#epicNoteInput").value = "";
  saveStore();
  renderAll();
}

function scoreExam() {
  let score = 0;
  examItems.forEach((item, index) => {
    const selected = document.querySelector(`input[name="exam-${index}"]:checked`)?.value;
    const note = document.querySelector(`#exam-note-${index}`);
    if (selected === item.answer) {
      score += 1;
      note.textContent = "Correct";
      note.className = "answer-note correct";
    } else {
      note.textContent = `Try again. The answer is ${item.answer}.`;
      note.className = "answer-note needs-practice";
      if (!activeProgress().weakSkills.includes(item.title.toLowerCase())) {
        activeProgress().weakSkills.unshift(item.title.toLowerCase());
      }
    }
  });
  activeProgress().examScores.push({ date: todayKey(), score, total: examItems.length });
  saveStore();
  renderParent();
}

function showWordPopover(word, target) {
  const key = normalizeWord(word);
  const entry = learningEntryForWord(key);
  const definition = entry.definition ?? "a word to learn by reading it in a sentence";
  const example = entry.example ?? `Try using ${key} in your own sentence.`;
  state.lastPopoverWord = key;
  const popover = document.querySelector("#wordPopover");
  document.querySelector("#popoverWord").textContent = entry.word || key;
  document.querySelector("#popoverDefinition").textContent = definition;
  document.querySelector("#popoverExample").textContent = example;
  popover.hidden = false;
  state.activePopoverAnchor = target ?? null;
  positionWordPopover();
}

function positionWordPopover() {
  const popover = document.querySelector("#wordPopover");
  if (!popover || popover.hidden) return;
  const margin = 14;
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  const rect = state.activePopoverAnchor?.getBoundingClientRect?.();
  const popoverRect = popover.getBoundingClientRect();
  const popoverWidth = Math.min(popoverRect.width || 380, viewportWidth - margin * 2);
  const popoverHeight = Math.min(popoverRect.height || 260, viewportHeight - margin * 2);
  const anchorVisible = rect && rect.bottom > margin && rect.top < viewportHeight - margin;
  const preferredLeft = anchorVisible ? rect.left : (viewportWidth - popoverWidth) / 2;
  const preferredTop = anchorVisible ? rect.bottom + 10 : Math.max(margin, (viewportHeight - popoverHeight) / 3);
  const left = Math.min(viewportWidth - popoverWidth - margin, Math.max(margin, preferredLeft));
  const top = Math.min(viewportHeight - popoverHeight - margin, Math.max(margin, preferredTop));
  popover.style.left = `${left}px`;
  popover.style.top = `${top}px`;
}

function preferredAudioMimeType() {
  const candidates = [
    "audio/mp4;codecs=mp4a.40.2",
    "audio/mp4",
    "audio/aac",
    "audio/webm;codecs=opus",
    "audio/webm"
  ];
  if (!window.MediaRecorder?.isTypeSupported) return "";
  return candidates.find((type) => MediaRecorder.isTypeSupported(type)) || "";
}

function createAudioRecorder(stream) {
  const mimeType = preferredAudioMimeType();
  state.activeRecordingMimeType = mimeType;
  return mimeType ? new MediaRecorder(stream, { mimeType }) : new MediaRecorder(stream);
}

function recordedAudioBlob() {
  const blobType = state.mediaRecorder?.mimeType || state.activeRecordingMimeType || "";
  return blobType
    ? new Blob(state.audioChunks, { type: blobType })
    : new Blob(state.audioChunks);
}

async function toggleRecording(button = document.querySelector("#recordButton")) {
  const promptIndex = Number(button?.dataset.recordPromptIndex ?? 0);
  const retell = buildWeeklyRetellPractice();
  const promptItem = normalizeRetellPrompt(retell.prompts[promptIndex] ?? retell.prompts[0], retell);
  if (!navigator.mediaDevices?.getUserMedia) {
    alert("Recording is not available in this browser.");
    return;
  }

  if (state.mediaRecorder?.state === "recording") {
    state.mediaRecorder.stop();
    const activeButton = state.activeRecordingButton ?? button;
    if (activeButton) activeButton.textContent = "Record again";
    state.activeRecordingButton = null;
    return;
  }

  const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
  state.audioChunks = [];
  state.activeRecordingButton = button;
  state.activeSpeakingRecordingIndex = promptIndex;
  state.retellParentChecked = { ...(state.retellParentChecked ?? {}), [promptIndex]: false };
  if (retell.sourceType === "raz-five-finger") {
    const razRetell = ensureAdventureDayProgress(activeProgress()).razRetell;
    razRetell.recordingDone = false;
    razRetell.replayDone = false;
    razRetell.parentChecked = false;
  }
  state.speakingReplayErrors = { ...(state.speakingReplayErrors ?? {}), [promptIndex]: "" };
  state.mediaRecorder = createAudioRecorder(stream);
  state.mediaRecorder.addEventListener("dataavailable", (event) => {
    if (event.data.size) state.audioChunks.push(event.data);
  });
  state.mediaRecorder.addEventListener("stop", () => {
    stream.getTracks().forEach((track) => track.stop());
    const blob = recordedAudioBlob();
    if (state.speakingRecordingUrls?.[promptIndex]) URL.revokeObjectURL(state.speakingRecordingUrls[promptIndex]);
    const url = URL.createObjectURL(blob);
    state.speakingRecordingUrls = { ...(state.speakingRecordingUrls ?? {}), [promptIndex]: url };
    state.activeSpeakingRecordingIndex = null;
    if (retell.sourceType === "raz-five-finger") {
      const razRetell = ensureAdventureDayProgress(activeProgress()).razRetell;
      razRetell.recordingDone = true;
      razRetell.replayDone = false;
      razRetell.parentChecked = false;
    }
    activeProgress().speakingAttempts.push({
      date: todayKey(),
      source: retell.sourceType === "raz-five-finger" ? "raz-five-finger-retell" : "weekly-reading-retell",
      retellSourceKey: retell.sourceKey,
      weeklyTitle: retell.weeklyTitle,
      target: promptItem.prompt,
      promptIndex,
      url,
      note: "recorded retell practice"
    });
    saveStore();
    renderPracticeRoute();
  });
  state.mediaRecorder.start();
  button.textContent = "Stop";
}

function replaySpeakingRecording(promptIndex = 0) {
  const url = state.speakingRecordingUrls?.[promptIndex];
  if (!url) return;
  state.speakingReplayErrors = { ...(state.speakingReplayErrors ?? {}), [promptIndex]: "" };
  const audio = new Audio(url);
  audio.play().then(() => {
    const retell = buildWeeklyRetellPractice();
    if (retell.sourceType === "raz-five-finger") {
      ensureAdventureDayProgress(activeProgress()).razRetell.replayDone = true;
      saveStore();
      renderPracticeRoute();
    }
  }).catch(() => {
    state.speakingReplayErrors = {
      ...(state.speakingReplayErrors ?? {}),
      [promptIndex]: "Replay could not play on this browser. Please record again."
    };
    renderPracticeRoute();
  });
}

function retellParentChecksFromProgress(retell = buildWeeklyRetellPractice(), progress = activeProgress()) {
  const checks = {};
  (progress.speakingAttempts ?? [])
    .filter((attempt) => attempt.date === todayKey()
      && attempt.source === "weekly-reading-retell"
      && (attempt.retellSourceKey ? attempt.retellSourceKey === retell.sourceKey : attempt.weeklyTitle === retell.weeklyTitle)
      && attempt.result === "parent checked")
    .forEach((attempt) => {
      const index = Number(attempt.promptIndex);
      if (Number.isInteger(index) && index >= 0) checks[index] = true;
    });
  return checks;
}

function completeRetellLevelIfReady(retell = buildWeeklyRetellPractice()) {
  if (retell.sourceType === "raz-five-finger") {
    const progress = ensureAdventureDayProgress(activeProgress()).razRetell;
    if (!progress.recordingDone || !progress.replayDone || !progress.parentChecked) return false;
    const speakCenter = missionCenters.find((item) => item.id === "speak");
    if (speakCenter && !centerProgress(speakCenter).complete) {
      recordPracticeAction("speak", { correct: true, items: 1, seconds: 60, render: false });
    }
    return true;
  }
  const promptItems = retell.prompts.map((prompt) => normalizeRetellPrompt(prompt, retell));
  const allPromptsChecked = promptItems.length > 0 && promptItems.every((_, index) => state.retellParentChecked?.[index]);
  if (!allPromptsChecked) return false;
  const speakCenter = missionCenters.find((item) => item.id === "speak");
  if (speakCenter && !centerProgress(speakCenter).complete) {
    recordPracticeAction("speak", { correct: true, items: 1, seconds: 60, render: false });
  }
  return true;
}

function completeRetellWithParent(promptIndex = 0) {
  const url = state.speakingRecordingUrls?.[promptIndex] || "";
  const retell = buildWeeklyRetellPractice();
  const promptItem = normalizeRetellPrompt(retell.prompts[promptIndex] ?? retell.prompts[0], retell);
  if (retell.sourceType === "raz-five-finger") {
    const razRetell = ensureAdventureDayProgress(activeProgress()).razRetell;
    if (!razRetell.recordingDone || !razRetell.replayDone) return;
    razRetell.parentChecked = true;
  }
  state.retellParentChecked = { ...(state.retellParentChecked ?? {}), [promptIndex]: true };
  activeProgress().speakingAttempts.push({
    date: todayKey(),
    source: retell.sourceType === "raz-five-finger" ? "raz-five-finger-retell" : "weekly-reading-retell",
    retellSourceKey: retell.sourceKey,
    weeklyTitle: retell.weeklyTitle,
    target: promptItem.prompt,
    promptIndex,
    url,
    result: "parent checked",
    feedback: "Checked with parents. Nice retell practice.",
    completedAt: new Date().toISOString()
  });
  completeRetellLevelIfReady(retell);
  saveStore();
  renderPracticeRoute();
}

function completeRetellPractice() {
  const retell = buildWeeklyRetellPractice();
  if (retell.missingStoryEvidence) {
    renderPracticeRoute();
    return;
  }
  const promptItems = retell.prompts.map((prompt) => normalizeRetellPrompt(prompt, retell));
  const allPromptsChecked = retell.sourceType === "raz-five-finger"
    ? completeRetellLevelIfReady(retell)
    : promptItems.length > 0 && promptItems.every((_, index) => state.retellParentChecked?.[index]);
  if (!allPromptsChecked) {
    renderPracticeRoute();
    return;
  }
  completeRetellLevelIfReady(retell);
  saveStore();
  returnToAdventureMap();
}

function handleStoryRewardAudioMetadata(event) {
  const audio = event.target;
  if (!(audio instanceof HTMLAudioElement) || !audio.matches("[data-story-reward-audio]")) return;
  const declared = Number(audio.dataset.declaredDuration);
  const actual = Number.isFinite(audio.duration) ? audio.duration : declared;
  if (Number.isFinite(actual) && actual > STORY_REWARD_MAX_AUDIO_SECONDS) {
    audio.pause();
    state.storyRewardAudioError = "This reward audio is longer than 5 minutes. Ask a parent to replace it with a shorter story.";
    renderPracticeRoute();
    return;
  }
  state.storyRewardAudioError = "";
}

async function toggleReadingWarmupRecording(button) {
  const razArticleId = button?.dataset.razArticleId || "";
  if (!navigator.mediaDevices?.getUserMedia) {
    alert("Recording is not available in this browser.");
    return;
  }

  if (state.mediaRecorder?.state === "recording" && state.readingWarmupRecordingActive) {
    state.mediaRecorder.stop();
    if (button) button.textContent = "Record again";
    return;
  }

  stopReadingWarmupReplay({ render: false });
  const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
  state.audioChunks = [];
  state.readingWarmupRecordingActive = true;
  state.readingWarmupActiveArticleId = razArticleId;
  state.readingWarmupReplayError = "";
  state.mediaRecorder = createAudioRecorder(stream);
  state.mediaRecorder.addEventListener("dataavailable", (event) => {
    if (event.data.size) state.audioChunks.push(event.data);
  });
  state.mediaRecorder.addEventListener("stop", () => {
    stream.getTracks().forEach((track) => track.stop());
    const blob = recordedAudioBlob();
    const recordingUrl = URL.createObjectURL(blob);
    if (razArticleId) {
      if (state.readingWarmupRecordingUrls?.[razArticleId]) URL.revokeObjectURL(state.readingWarmupRecordingUrls[razArticleId]);
      state.readingWarmupRecordingUrls = { ...(state.readingWarmupRecordingUrls ?? {}), [razArticleId]: recordingUrl };
    } else {
      if (state.readingWarmupRecordingUrl) URL.revokeObjectURL(state.readingWarmupRecordingUrl);
      state.readingWarmupRecordingUrl = recordingUrl;
    }
    state.readingWarmupRecordingActive = false;
    state.readingWarmupActiveArticleId = "";
    const dayProgress = ensureAdventureDayProgress(activeProgress());
    if (razArticleId) {
      dayProgress.razReadingArticles[razArticleId] = {
        ...(dayProgress.razReadingArticles[razArticleId] ?? {}),
        recordingDone: true,
        replayDone: false
      };
      syncRazReadingAggregate(dayProgress);
    } else {
      dayProgress.readingWarmupRecordingDone = true;
      dayProgress.readingWarmupReplayDone = false;
    }
    saveStore();
    renderPracticeRoute();
  });
  state.mediaRecorder.start();
  if (button) button.textContent = "Stop";
}

function stopReadingWarmupReplay({ render = true } = {}) {
  const audio = state.readingWarmupReplayAudio;
  if (audio) {
    audio.pause();
    try {
      audio.currentTime = 0;
    } catch {
      // Some mobile browsers do not allow seeking blob audio after teardown.
    }
  }
  state.readingWarmupReplayAudio = null;
  state.readingWarmupReplayActive = false;
  state.readingWarmupReplayArticleId = "";
  if (render) renderPracticeRoute();
}

function replayReadingWarmupRecording(articleId = "") {
  const recordingUrl = articleId ? state.readingWarmupRecordingUrls?.[articleId] : state.readingWarmupRecordingUrl;
  if (!recordingUrl) return;
  if (state.readingWarmupReplayActive) {
    stopReadingWarmupReplay();
    return;
  }
  state.readingWarmupReplayError = "";
  const audio = new Audio(recordingUrl);
  const isCurrentReplay = () => state.readingWarmupReplayAudio === audio;
  state.readingWarmupReplayAudio = audio;
  state.readingWarmupReplayActive = true;
  state.readingWarmupReplayArticleId = articleId;
  audio.addEventListener("ended", () => {
    if (!isCurrentReplay()) return;
    stopReadingWarmupReplay();
  }, { once: true });
  audio.addEventListener("error", () => {
    if (!isCurrentReplay()) return;
    stopReadingWarmupReplay({ render: false });
    state.readingWarmupReplayError = "Replay could not play on this browser. Please record again.";
    renderPracticeRoute();
  }, { once: true });
  audio.play().then(() => {
    if (!isCurrentReplay()) return;
    const dayProgress = ensureAdventureDayProgress(activeProgress());
    if (articleId) {
      dayProgress.razReadingArticles[articleId] = {
        ...(dayProgress.razReadingArticles[articleId] ?? {}),
        replayDone: true
      };
      syncRazReadingAggregate(dayProgress);
    } else {
      dayProgress.readingWarmupReplayDone = true;
    }
    saveStore();
    renderPracticeRoute();
  }).catch(() => {
    if (!isCurrentReplay()) return;
    stopReadingWarmupReplay({ render: false });
    state.readingWarmupReplayError = "Replay could not play on this browser. Please record again.";
    renderPracticeRoute();
  });
  renderPracticeRoute();
}

function evaluateSpeakingTranscript(transcript, prompt) {
  const retell = buildWeeklyRetellPractice();
  const normalized = transcript.toLowerCase().replace(/[^a-z' ]/g, " ");
  const words = normalized.split(/\s+/).filter(Boolean);
  const keywordMatches = retell.keywords
    .map((word) => normalizeWord(word))
    .filter((word) => word && normalized.includes(word));
  const retellSignals = retell.retellSignals ?? ["first", "next", "then", "because", "problem", "solution", "learns", "changes"];
  const signalMatches = retellSignals.filter((word) => normalized.includes(word));
  const strong = words.length >= 10 && keywordMatches.length >= 2 && signalMatches.length >= 1;
  const okay = words.length >= 5 && (keywordMatches.length >= 1 || signalMatches.length >= 1);
  if (strong) {
    return {
      result: "strong retell",
      feedback: "Strong retell. You used story details and weekly words.",
      correct: true,
      keywordMatches,
      weeklyTitle: retell.weeklyTitle,
      prompt
    };
  }
  if (okay) {
    return {
      result: "building retell",
      feedback: "Good start. Add one more story detail or weekly word next time.",
      correct: true,
      keywordMatches,
      weeklyTitle: retell.weeklyTitle,
      prompt
    };
  }
  return {
    result: "needs more detail",
    feedback: "Try again with a story detail and one weekly word.",
    correct: false,
    keywordMatches,
    weeklyTitle: retell.weeklyTitle,
    prompt
  };
}

function startSpeechCheck(promptIndex = 0) {
  const retell = buildWeeklyRetellPractice();
  const prompt = retell.prompts[promptIndex] ?? retell.prompts[0];
  state.speakingTarget = prompt;
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    alert("Speech recognition is not available in this browser. You can still record and replay.");
    return;
  }

  const recognition = new SpeechRecognition();
  recognition.lang = "en-US";
  recognition.interimResults = false;
  recognition.maxAlternatives = 1;
  recognition.addEventListener("result", (event) => {
    const transcript = event.results[0][0].transcript.toLowerCase();
    const evaluation = evaluateSpeakingTranscript(transcript, prompt);
    activeProgress().speakingAttempts.push({
      date: todayKey(),
      source: "weekly-reading-retell",
      weeklyTitle: retell.weeklyTitle,
      target: prompt,
      transcript,
      result: evaluation.result,
      feedback: evaluation.feedback,
      keywordMatches: evaluation.keywordMatches,
      completedAt: new Date().toISOString()
    });
    if (!evaluation.correct && !activeProgress().weakSkills.includes("speaking fluency")) {
      activeProgress().weakSkills.unshift("speaking fluency");
    }
    recordPracticeAction("speak", { correct: evaluation.correct, items: 1, seconds: 60, render: false });
    saveStore();
    alert(`${evaluation.feedback}\nI heard: "${transcript}"`);
    renderAll();
  });
  recognition.addEventListener("error", () => {
    alert("Speech check could not hear clearly. Try recording and replaying instead.");
  });
  recognition.start();
}

document.addEventListener("click", (event) => {
  const startProfile = event.target.closest("[data-start-profile]");
  if (startProfile) {
    selectProfile(startProfile.dataset.startProfile);
    return;
  }

  if (event.target.closest("[data-close-heart-reward]")) {
    state.heartRewardOverlay = null;
    renderHeartRewardOverlay();
    return;
  }

  if (event.target.closest("[data-cloud-sync-now]")) {
    updateCloudSyncStatus("syncing", "Syncing now...");
    hydrateStoreFromCloud();
    return;
  }

  if (event.target.closest("[data-open-parent-auth]")) {
    showParentAuth();
    return;
  }

  const openLearner = event.target.closest("[data-open-learner]");
  if (openLearner) {
    selectProfile(openLearner.dataset.openLearner);
    return;
  }

  const editName = event.target.closest("[data-edit-name]");
  if (editName) {
    state.editingProfileId = editName.dataset.editName;
    renderProfileGate();
    setTimeout(() => document.querySelector(`[data-profile-name="${state.editingProfileId}"]`)?.focus(), 0);
    return;
  }

  const finishName = event.target.closest("[data-finish-name]");
  if (finishName) {
    state.editingProfileId = null;
    renderProfileGate();
    return;
  }

  const learnerButton = event.target.closest("[data-learner-id]");
  if (learnerButton) {
    stopAudio();
    state.activeLearnerId = learnerButton.dataset.learnerId;
    renderAll();
    return;
  }

  const navButton = event.target.closest("[data-section]");
  if (navButton) {
    stopAudio();
    state.activeSection = navButton.dataset.section;
    renderAll();
    if (state.activeSection === "materials") loadMaterials();
    return;
  }

  const audioButtonTarget = event.target.closest("[data-audio-target]");
  if (audioButtonTarget) {
    playAudio(audioButtonTarget.dataset.audioTarget, audioButtonTarget.dataset.audioKey, audioButtonTarget);
    return;
  }

  const spellingAudio = event.target.closest("[data-spelling-audio]");
  if (spellingAudio) {
    speakSpellingPrompt(spellingAudio.dataset.spellingAudio, spellingAudio.dataset.audioKey, spellingAudio);
    return;
  }

  const readingWarmupRecord = event.target.closest("[data-reading-warmup-record]");
  if (readingWarmupRecord) {
    toggleReadingWarmupRecording(readingWarmupRecord).catch(() => alert("Microphone permission is needed for recording."));
    return;
  }

  const readingWarmupReplay = event.target.closest("[data-reading-warmup-replay]");
  if (readingWarmupReplay) {
    replayReadingWarmupRecording(readingWarmupReplay.dataset.razArticleId || "");
    return;
  }

  const razArticleSelect = event.target.closest("[data-raz-article-select]");
  if (razArticleSelect) {
    if (state.readingWarmupRecordingActive) return;
    stopReadingWarmupReplay({ render: false });
    state.razActiveArticleId = razArticleSelect.dataset.razArticleSelect;
    renderPracticeRoute();
    return;
  }

  if (event.target.closest("[data-close-raz-inventory]")) {
    state.razInventoryPopup = false;
    returnToAdventureMap();
    return;
  }

  if (event.target.closest("[data-reading-warmup-complete]")) {
    const result = completeReadingWarmup({ render: false });
    if (result?.completed) {
      if (isSummerRazDate() && razInventoryIsLow()) {
        state.razInventoryPopup = true;
        renderPracticeRoute();
      } else {
        returnToAdventureMap();
      }
    } else {
      renderPracticeRoute();
    }
    return;
  }

  if (event.target.closest("[data-reading-warmup-skip]")) {
    const result = skipReadingWarmup({ render: false });
    if (result?.completed) {
      returnToAdventureMap();
    } else {
      renderPracticeRoute();
    }
    return;
  }

  const wordPowerWord = event.target.closest("[data-word-power-word]");
  if (wordPowerWord) {
    recordWordPowerWord(wordPowerWord.dataset.wordPowerWord, { render: false });
    showWordPopover(wordPowerWord.dataset.wordPowerWord, wordPowerWord);
    renderPracticeRoute();
    return;
  }

  if (event.target.closest("[data-complete-word-power]")) {
    const result = completeWordPower({ render: false });
    if (result?.completed || result?.complete) {
      returnToAdventureMap();
    } else {
      renderPracticeRoute();
    }
    return;
  }

  if (event.target.closest("[data-finish-story-reward]")) {
    completeStoryReward({ render: false });
    returnToAdventureMap();
    return;
  }

  if (event.target.closest("[data-grammar-reward-ready]")) {
    returnToAdventureMap();
    return;
  }

  const startPractice = event.target.closest("[data-start-practice]");
  if (startPractice) {
    stopAudio();
    state.activePracticeId = startPractice.dataset.startPractice;
    state.practiceMode = "route";
    openPracticeRoute(startPractice.dataset.startPractice);
    return;
  }

  if (event.target.id === "backToTodayButton") {
    stopAudio();
    stopReadingWarmupReplay({ render: false });
    const route = currentPracticeRoute();
    const parentRoute = route ? parentPracticeRoute(route.practiceId) : "";
    if (parentRoute) {
      openPracticeRoute(parentRoute);
    } else if (route) {
      state.practiceMode = "overview";
      state.activePracticeId = null;
      window.location.hash = "";
    } else {
      state.practiceMode = "overview";
      state.activePracticeId = null;
      renderToday();
    }
    return;
  }

  const practiceAction = event.target.closest("[data-practice-action]");
  if (practiceAction) {
    const actionMap = {
      sight: "sight",
      wordwork: "wordwork",
      reading: "reading",
      writing: "writing",
      warmup: "warmup"
    };
    const centerId = actionMap[practiceAction.dataset.practiceAction];
    if (centerId) recordPracticeAction(centerId, { correct: true, items: 1, seconds: 30 });
    return;
  }

  if (event.target.closest("[data-sight-choice], [data-sight-fill], [data-sight-build-letter], [data-sight-build-clear], [data-sight-hunt-word], [data-sight-read], [data-sight-sentence-match]")) {
    handleSightAction(event.target);
    return;
  }

  if (event.target.closest("[data-reading-quiz-choice]")) {
    handleReadingQuizChoice(event.target);
    return;
  }

  if (event.target.closest("[data-reading-quiz-next]")) {
    handleReadingQuizNext();
    return;
  }

  if (event.target.closest("[data-grammar-choice]")) {
    handleGrammarChoice(event.target);
    return;
  }

  const wordResult = event.target.closest("[data-word-result]");
  if (wordResult) {
    const [word, result] = wordResult.dataset.wordResult.split(":");
    scheduleWordReview(word, result === "correct" ? "correct" : "practice");
    return;
  }

  const wordLink = event.target.closest("[data-word]");
  if (wordLink) {
    if (wordLink.classList.contains("vocab-word-button")) {
      recordWordPowerWord(wordLink.dataset.word, { render: false });
    }
    showWordPopover(wordLink.dataset.word, wordLink);
    return;
  }

  if (event.target.id === "closePopoverButton") {
    stopAudio();
    document.querySelector("#wordPopover").hidden = true;
    state.activePopoverAnchor = null;
    return;
  }

  if (event.target.id === "speakWordButton") {
    const word = document.querySelector("#popoverWord")?.textContent ?? state.lastPopoverWord;
    speakWordHelp(word, event.target, "popover");
    return;
  }

  if (event.target.closest("[data-close-spelling-feedback]")) {
    if (state.spellingFeedbackOverlay?.type === "saved") {
      advanceSpellingWord(0);
    } else {
      state.spellingFeedbackOverlay = null;
      renderPracticeRoute();
    }
    return;
  }

  const recordPromptButton = event.target.closest("[data-record-prompt-index], #recordButton");
  if (recordPromptButton) {
    toggleRecording(recordPromptButton).catch(() => alert("Microphone permission is needed for recording."));
    return;
  }

  const razRetellSelect = event.target.closest("[data-raz-retell-select]");
  if (razRetellSelect) {
    const articleId = razRetellSelect.dataset.razRetellSelect;
    const dayProgress = ensureAdventureDayProgress(activeProgress());
    if (!dayProgress.razReadingAssignments.includes(articleId)) return;
    Object.values(state.speakingRecordingUrls ?? {}).forEach((url) => URL.revokeObjectURL(url));
    state.speakingRecordingUrls = {};
    state.speakingReplayErrors = {};
    state.retellParentChecked = {};
    dayProgress.razRetell = {
      ...defaultRazRetellProgress(),
      selectedArticleId: articleId,
      selectedAt: new Date().toISOString()
    };
    saveStore();
    renderPracticeRoute();
    return;
  }

  if (event.target.closest("[data-change-raz-retell-book]")) {
    if (state.mediaRecorder?.state === "recording") return;
    const dayProgress = ensureAdventureDayProgress(activeProgress());
    dayProgress.razRetell = defaultRazRetellProgress();
    state.retellSessionKey = "";
    saveStore();
    renderPracticeRoute();
    return;
  }

  const retellReplayButton = event.target.closest("[data-retell-replay]");
  if (retellReplayButton) {
    replaySpeakingRecording(Number(retellReplayButton.dataset.retellReplay ?? 0));
    return;
  }

  const retellParentCheckButton = event.target.closest("[data-retell-parent-check]");
  if (retellParentCheckButton) {
    completeRetellWithParent(Number(retellParentCheckButton.dataset.retellParentCheck ?? 0));
    return;
  }

  if (event.target.closest("[data-complete-retell-practice]")) {
    completeRetellPractice();
    return;
  }

  const speechCheckButton = event.target.closest("[data-speech-check], #speechCheckButton");
  if (speechCheckButton) {
    startSpeechCheck(Number(speechCheckButton.dataset.speakPromptIndex ?? 0));
    return;
  }

  if (event.target.id === "parentLogoutButton") {
    logoutParent();
    return;
  }

  if (event.target.id === "archiveThisWeekButton") {
    const notes = window.prompt("Parent notes for this weekly archive:", "") || "";
    const archive = archiveCurrentWeek(notes);
    state.materialsStatus = `${archive.archiveLabel} archived for parent meetings.`;
    return;
  }

  if (event.target.id === "clearAllTestProgressButton") {
    clearAllTestPracticeRecords();
    return;
  }

  if (event.target.closest("[data-create-weekly-plan]")) {
    createWeeklyPlanFromBatch();
    return;
  }

  if (event.target.closest("[data-dismiss-material-batch]")) {
    state.materialBatch = null;
    state.materialsStatus = "";
    renderMaterials();
    return;
  }

  if (event.target.closest("[data-clear-uploaded-materials]")) {
    clearCurrentUploadedWeeklyPlanMaterials();
    return;
  }

  const materialButton = event.target.closest("[data-open-material]");
  if (materialButton) {
    openMaterial(materialButton.dataset.openMaterial);
    return;
  }

  const useMaterialButton = event.target.closest("[data-use-material]");
  if (useMaterialButton) {
    useMaterialForWeek(useMaterialButton.dataset.useMaterial);
    return;
  }

  const deleteMaterialButton = event.target.closest("[data-delete-material]");
  if (deleteMaterialButton) {
    deleteMaterialAndTrainingPlans(deleteMaterialButton.dataset.deleteMaterial);
    return;
  }

  const applyMaterialSuggestionsButton = event.target.closest("[data-apply-material-suggestions], #applyMaterialSuggestionsButton");
  if (applyMaterialSuggestionsButton) {
    syncWeeklySuggestionEditsFromReview();
    applyWeeklySuggestions(applyMaterialSuggestionsButton.dataset.applyMaterialSuggestions || "auto");
    return;
  }

  if (event.target.id === "dismissMaterialSuggestionsButton") {
    state.weeklySuggestions = null;
    state.materialBatch = null;
    renderMaterials();
  }
});

document.addEventListener("submit", (event) => {
  if (event.target.id === "spellingAnswerForm") {
    handleSpellingSubmit(event);
    return;
  }

  if (event.target.id === "parentAuthForm") {
    unlockParent(event);
    return;
  }

  if (event.target.id === "materialUploadForm") {
    uploadMaterial(event);
    return;
  }

  if (event.target.id === "parentPasswordForm") {
    changeParentPassword(event);
    return;
  }

  if (event.target.id !== "epicShelfForm") return;
  event.preventDefault();
  const title = document.querySelector("#shelfTitleInput").value.trim();
  const tags = uniqueWords(document.querySelector("#shelfTagsInput").value.split(/,|\n/).map((tag) => tag.toLowerCase()));
  if (!title) return;
  store.parentBooks.unshift({
    title,
    author: "Parent added",
    source: "Parent shelf",
    sourceUrl: "https://www.getepic.com/",
    gradeBand: "G1-G2",
    format: tags.includes("fiction") ? "Fiction" : tags.includes("decodable") ? "Decodable" : "Epic book",
    readTime: "10-15 min",
    topicTags: tags,
    skillTags: tags,
    searchTerms: [title, ...tags]
  });
  saveStore();
  renderAll();
});

document.addEventListener("input", (event) => {
  if (event.target.id === "spellingAnswerInput") {
    handleSpellingInput(event);
    return;
  }

  if (event.target.id === "materialSearchInput") {
    state.materialSearch = event.target.value;
    renderMaterials();
    return;
  }

  const input = event.target.closest("[data-profile-name]");
  if (!input) return;
  const fallback = profiles.find((profile) => profile.id === input.dataset.profileName)?.defaultName ?? "Learner";
  store.profileNames[input.dataset.profileName] = input.value.trim() || fallback;
  saveStore();
  const button = input.closest(".profile-card")?.querySelector("[data-start-profile]");
  if (button) button.textContent = store.profileNames[input.dataset.profileName];
});

document.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  const input = event.target.closest("[data-profile-name]");
  if (!input) return;
  event.preventDefault();
  state.editingProfileId = null;
  renderProfileGate();
});

document.querySelector("#weeklyForm").addEventListener("submit", saveWeeklyPlan);
document.querySelector("#epicForm").addEventListener("submit", logEpic);
document.querySelector("#editWeeklyButton").addEventListener("click", () => {
  state.weeklyPlanEditing = true;
  renderWeeklyForm();
  document.querySelector("#readingTitleInput")?.focus();
});
document.querySelector("#resetWeeklyButton").addEventListener("click", () => {
  if (!state.weeklyPlanEditing) return;
  const weekInfo = learningWeekInfo();
  store.weeklyPlan = applyWeekMetadata({ ...defaultWeeklyPlan, updatedAt: new Date().toISOString() }, weekInfo, "active");
  store.currentWeekId = store.weeklyPlan.weekId;
  if (!store.weeklyPlansByWeekId) store.weeklyPlansByWeekId = {};
  store.weeklyPlansByWeekId[weekInfo.weekId] = store.weeklyPlan;
  saveStore();
  renderAll();
});
document.querySelector("#startFridayButton").addEventListener("click", () => {
  const word = activePlan().masteryWords[0];
  if (word) playAudio(word, `friday-start-${normalizeWord(word)}`, document.querySelector("#startFridayButton"));
});
document.querySelector("#markReviewsDoneButton").addEventListener("click", () => {
  dueReviewWords().forEach((item) => scheduleWordReview(item.word, "correct"));
});
document.querySelector("#practiceDomainSelect").addEventListener("change", (event) => {
  state.practiceDomain = event.target.value;
  renderPractice();
});
document.querySelector("#scoreExamButton").addEventListener("click", scoreExam);
document.querySelector("#clearProgressButton").addEventListener("click", () => {
  store.learners[state.activeLearnerId] = getDefaultProgress();
  saveStore();
  renderAll();
});
document.querySelector("#backToProfilesButton").addEventListener("click", showProfileGate);
document.querySelector("#childSwitchButton").addEventListener("click", showProfileGate);
document.querySelector("#cancelParentAuthButton").addEventListener("click", hideParentAuth);
document.querySelector("#refreshMaterialsButton").addEventListener("click", loadMaterials);
document.addEventListener("loadedmetadata", handleStoryRewardAudioMetadata, true);
window.addEventListener("hashchange", syncPracticeRoute);
window.addEventListener("resize", positionWordPopover);
window.addEventListener("scroll", positionWordPopover, true);
Object.assign(window, {
  completeAdventureLevel,
  isAdventureLevelUnlocked,
  getTodayAdventureLevels,
  recordWordPowerWord,
  completeWordPower,
  completeReadingWarmup,
  skipReadingWarmup,
  readingWarmupSkipStatus,
  completeStoryReward,
  getWordPowerStatus
});
seedSightWordBank();
activateWeeklyPlanForToday();
ensureCw25ManualWeeklyPlanFix();
ensureCw25StoryRewardAudio();
ensureCw26StoryRewardAudio();
ensureCw26Vocabulary();
ensureCurrentWeekArchiveSeed();
resetTesterForCurrentBuild();
setupCloudSync();

if (currentPracticeRoute()) {
  syncPracticeRoute();
} else {
  showProfileGate();
}
