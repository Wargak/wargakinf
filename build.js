const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  BorderStyle, LevelFormat, Numbering, convertInchesToTwip
} = require("docx");
const fs = require("fs");

const ACCENT = "1F4E5F"; // deep teal
const GREY = "555555";

const numbering = {
  config: [
    {
      reference: "bullets",
      levels: [
        {
          level: 0,
          format: LevelFormat.BULLET,
          text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: convertInchesToTwip(0.28), hanging: convertInchesToTwip(0.18) } } },
        },
      ],
    },
  ],
};

function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 320, after: 120 },
    border: { bottom: { color: ACCENT, space: 4, style: BorderStyle.SINGLE, size: 8 } },
    children: [
      new TextRun({ text: text.toUpperCase(), bold: true, color: ACCENT, size: 22, font: "Calibri" }),
    ],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 40 },
    children: [new TextRun({ text, size: 21, font: "Calibri" })],
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 80 },
    children: [new TextRun({ text, size: 21, font: "Calibri", italics: opts.italics, bold: opts.bold })],
  });
}

function subHeading(text) {
  return new Paragraph({
    spacing: { before: 200, after: 60 },
    children: [new TextRun({ text, bold: true, size: 21, font: "Calibri", color: "222222" })],
  });
}

const skillGroups = [
  { label: "Languages", items: ["HTML5", "CSS3", "JavaScript", "Python", "C++", "SQL"] },
  { label: "Frameworks & Technologies", items: ["React", "Firebase", "SQLite", "Responsive Web Technologies"] },
  { label: "Development Tools", items: ["GitHub", "VS Code", "Android development/testing environments", "Browser Developer Tools"] },
];

const coreSkills = [
  "Frontend Web Development", "Responsive Web Design", "UI/UX Design",
  "Mobile Application Development", "Mobile Phone Hardware & Software Solutions",
  "Website Design & Development", "Interactive User Interfaces", "Database Fundamentals",
  "Software Troubleshooting", "Graphic & Digital Design", "GitHub & Website Deployment",
  "Problem Solving", "Creative Thinking", "Self-Directed Learning",
];

const projects = [
  {
    title: "Personal Portfolio Website",
    desc: "Designed and developed a professional personal portfolio concept showcasing skills, services, projects, and professional identity.",
    features: ["Responsive design", "Modern UI/UX", "Project showcase", "Skills and services sections", "Contact section", "Interactive navigation"],
  },
  {
    title: "Nuer Gospel Songs & Lyrics Platform",
    desc: "Website/application concept allowing users to read gospel lyrics and listen to songs, focused on preserving and making Nuer gospel music more accessible digitally.",
    features: ["Lyrics database", "Artist profiles", "Song categories", "Audio integration", "Search functionality", "Responsive mobile interface"],
  },
  {
    title: "CM Language Application",
    desc: "Concept application supporting Nuer, English, and Amharic content for reading and language accessibility.",
    features: [],
  },
  {
    title: "Wargak Mobile Repair System",
    desc: "Concept for a mobile repair management platform.",
    features: ["Customer registration", "Sign-in/sign-up", "Phone repair requests", "Problem descriptions", "Repair status updates", "SMS/email notification concepts", "Customer dashboard", "Admin panel", "Service management"],
  },
  {
    title: "7 Light Music Studio Platform",
    desc: "Music studio website concept with modern interfaces for artists and music producers.",
    features: ["User authentication", "Dashboard", "Artist dashboard", "Music-related content", "Interactive animations", "Modern branding and visual design"],
  },
  {
    title: "E-Commerce Website",
    desc: "Selling platform concept.",
    features: ["User registration", "Sign-in", "Product listings", "Shopping cart", "User dashboard", "Admin panel"],
  },
  {
    title: "Gallery Application",
    desc: "Responsive gallery application concept using web technologies with a focus on mobile-friendly presentation and interaction.",
    features: [],
  },
  {
    title: "Phone Banking System",
    desc: "Python/C++ programming project involving banking-system concepts and database-related functionality.",
    features: [],
  },
];

const children = [];

// Header
children.push(new Paragraph({
  spacing: { after: 20 },
  children: [new TextRun({ text: "WARGAK NHIAL LOGOTE", bold: true, size: 40, color: ACCENT, font: "Calibri" })],
}));
children.push(new Paragraph({
  spacing: { after: 140 },
  children: [new TextRun({ text: "Web Developer | UI/UX Designer | Mobile Technology Specialist", size: 22, color: GREY, font: "Calibri" })],
}));
children.push(new Paragraph({
  border: { bottom: { color: ACCENT, space: 6, style: BorderStyle.SINGLE, size: 12 } },
  spacing: { after: 160 },
  children: [
    new TextRun({ text: "Ethiopia   |   napsy.atomn06@gmail.com   |   +251 980 855 560", size: 20, color: GREY, font: "Calibri" }),
  ],
}));

// Summary
children.push(sectionHeading("Professional Summary"));
children.push(body("Motivated and creative technology professional with hands-on experience in frontend web development, UI/UX design, mobile software and hardware solutions, and digital product development. Skilled in building responsive and interactive websites and developing practical applications using modern technologies."));
children.push(body("Passionate about creating clean, user-friendly digital experiences and continuously expanding technical knowledge toward full-stack development. Experienced in working on personal and community-focused projects, including websites, mobile applications, dashboards, music platforms, educational systems, and digital content platforms."));

// Core Skills
children.push(sectionHeading("Core Skills"));
coreSkills.forEach(s => children.push(bullet(s)));

// Technical Skills
children.push(sectionHeading("Technical Skills"));
skillGroups.forEach(g => {
  children.push(subHeading(g.label));
  children.push(body(g.items.join("  \u2022  ")));
});

// Project Experience
children.push(sectionHeading("Project Experience"));
projects.forEach(p => {
  children.push(subHeading(p.title));
  children.push(body(p.desc, { after: 60 }));
  if (p.features.length) {
    children.push(body("Key features:", { italics: true, after: 40 }));
    p.features.forEach(f => children.push(bullet(f)));
  }
});

// UI/UX & Design
children.push(sectionHeading("UI/UX & Design"));
children.push(body("Strong interest in creating interfaces that are:", { after: 60 }));
["Clean and modern", "Easy to navigate", "Mobile responsive", "User friendly", "Visually engaging", "Accessible and practical"].forEach(f => children.push(bullet(f)));
children.push(body("Enjoys designing interfaces for websites, mobile applications, dashboards, portfolios, business platforms, and digital products.", { after: 60 }));

// Mobile Technology
children.push(sectionHeading("Mobile Technology"));
children.push(body("Hands-on experience and interest in:", { after: 60 }));
["Mobile phone troubleshooting", "Mobile software solutions", "Mobile hardware solutions", "Application testing", "Android environments", "Mobile-responsive website development", "Digital device problem solving"].forEach(f => children.push(bullet(f)));

// Current Development Goals
children.push(sectionHeading("Current Development Goals"));
children.push(body("Currently expanding knowledge toward becoming a full-stack developer, with particular interest in:", { after: 60 }));
["Advanced JavaScript", "React", "Backend development", "Python", "Databases", "Firebase", "APIs", "Software architecture", "Modern web application development"].forEach(f => children.push(bullet(f)));

// Professional Strengths
children.push(sectionHeading("Professional Strengths"));
[
  ["Problem Solver", "Enjoys identifying technical problems and finding practical solutions."],
  ["Creative Thinker", "Combines technology and design to create useful and attractive digital products."],
  ["Self-Motivated Learner", "Continuously learns programming and technology through practical projects and independent study."],
  ["Adaptable", "Comfortable exploring different technologies and applying them to new projects."],
  ["User-Focused", "Pays attention to usability, responsiveness, and the overall user experience."],
].forEach(([label, desc]) => {
  children.push(new Paragraph({
    spacing: { after: 60 },
    children: [
      new TextRun({ text: label + ": ", bold: true, size: 21, font: "Calibri" }),
      new TextRun({ text: desc, size: 21, font: "Calibri" }),
    ],
  }));
});

// Services
children.push(sectionHeading("Services"));
["Website Development", "Frontend Development", "Responsive Web Design", "UI/UX Design", "Mobile Website Development", "Mobile Software Support", "Mobile Hardware Solutions", "Website Maintenance", "Digital Product Design", "Basic Database Development"].forEach(f => children.push(bullet(f)));

// Languages
children.push(sectionHeading("Languages"));
children.push(body("English  \u2022  Amharic  \u2022  Nuer"));

// Career Objective
children.push(sectionHeading("Career Objective"));
children.push(body("To work with a forward-thinking organization, company, startup, or development team where I can apply my skills in web development, UI/UX design, mobile technology, and software development, while continuing to grow as a professional full-stack developer and contributing to meaningful technology projects."));

// References
children.push(sectionHeading("References"));
children.push(body("Available upon request."));

const doc = new Document({
  numbering,
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 21 } },
    },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 }, // US Letter
          margin: { top: 720, bottom: 720, left: 900, right: 900 },
        },
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/home/claude/cv/Wargak_Nhial_Logote_CV.docx", buf);
  console.log("done");
});
