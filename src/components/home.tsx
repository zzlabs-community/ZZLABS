"use client";
import Image from "next/image";
import {motion} from "motion/react";
import {ArrowUpRight, Code2, Globe2, Sparkles, Workflow, Layers3, Menu, X, Boxes, MessageCircle} from "lucide-react";
import {useEffect,useState} from "react";
import RippleGrid from "./RippleGrid";

import TrueFocus from "./TrueFocus";
import ShinyText from "./ShinyText";
import StarBorder from "./StarBorder";

type Lang="es"|"en";
const C={
es:{
nav:["Sistemas","Servicios","Proyectos","Fundador","Contacto"],
heroEye:"Tecnología que convierte ideas en sistemas reales.",
hero1:"Construimos sistemas",hero2:"que hacen avanzar.",
heroP:"Diseñamos productos digitales, plataformas SaaS y automatizaciones pensadas para operar de verdad: rápidas, escalables y listas para crecer.",
explore:"Explorar ZZLabs",
stat1:"Producto activo",stat2:"Tecnologías",stat3:"Enfoque",
stats:["Nexo Ministerial","Full-stack + IA","Sistemas escalables"],
systems:"SISTEMAS",products:"Tecnología convertida en producto.",
nexo:"Plataforma integral para la gestión de iglesias: personas, asistencia, ministerios, células, seguimiento, economía, reportes y administración multi-iglesia.",
view:"Ver proyecto",
services:"QUÉ HACEMOS",servicesTitle:"De la idea a un sistema que funciona.",
serviceDesc:["Productos web y SaaS construidos con una arquitectura lista para crecer.","Automatizaciones para reducir trabajo manual y conectar procesos reales.","Integraciones de IA aplicadas donde generan valor medible."],
focus:"NUESTRO ENFOQUE",focusSentence:"Diseñar Construir Escalar",
tech:"STACK",techTitle:"Tecnología moderna, sin adornos innecesarios.",
projects:"PROYECTOS DE GITHUB",projectsTitle:"El laboratorio sigue creciendo.",
projectsP:"Esta galería se alimenta automáticamente de los repositorios públicos de ZZLabs. Cada nuevo proyecto publicado en GitHub aparece aquí con su nombre y tecnología.",
githubCta:"Ver todos en GitHub",
process:"PROCESO",processTitle:"Menos presentación. Más construcción.",
steps:[["01","Entender","Definimos el problema, los usuarios y qué resultado debe producir el sistema."],["02","Arquitectar","Elegimos una base técnica que pueda crecer sin rehacer todo después."],["03","Construir","Desarrollamos, probamos y refinamos sobre producto real."],["04","Evolucionar","Medimos, automatizamos y agregamos capacidad cuando el negocio lo necesita."]],
founder:"FUNDADOR",founderRole:"Fundador & Arquitecto de Sistemas",founderP:"Desarrollador Full Stack enfocado en productos digitales, automatización y sistemas que conectan tecnología con operaciones reales.",
contact:"TRABAJEMOS",contact1:"¿Tienes una idea que",contact2:"merece existir?",talk:"Hablemos"
},
en:{
nav:["Systems","Services","Projects","Founder","Contact"],
heroEye:"Technology that turns ideas into real systems.",
hero1:"We build systems",hero2:"that move things forward.",
heroP:"We design digital products, SaaS platforms and automation built for real operations: fast, scalable and ready to grow.",
explore:"Explore ZZLabs",
stat1:"Active product",stat2:"Technology",stat3:"Focus",
stats:["Nexo Ministerial","Full-stack + AI","Scalable systems"],
systems:"SYSTEMS",products:"Technology turned into product.",
nexo:"An end-to-end church management platform: people, attendance, ministries, groups, follow-up, finance, reporting and multi-church administration.",
view:"View project",
services:"WHAT WE DO",servicesTitle:"From idea to a system that works.",
serviceDesc:["Web products and SaaS built on architecture ready to grow.","Automation that reduces manual work and connects real processes.","AI integrations applied where they create measurable value."],
focus:"OUR FOCUS",focusSentence:"Design Build Scale",
tech:"STACK",techTitle:"Modern technology, without unnecessary decoration.",
projects:"GITHUB PROJECTS",projectsTitle:"The lab keeps growing.",
projectsP:"This gallery is automatically fed by ZZLabs public repositories. Every new GitHub project appears here with its name and technology.",
githubCta:"See all on GitHub",
process:"PROCESS",processTitle:"Less presentation. More building.",
steps:[["01","Understand","Define the problem, users and the result the system needs to produce."],["02","Architect","Choose a technical foundation that can grow without rebuilding everything later."],["03","Build","Develop, test and refine against a real product."],["04","Evolve","Measure, automate and add capability when the business needs it."]],
founder:"FOUNDER",founderRole:"Founder & Systems Architect",founderP:"Full Stack developer focused on digital products, automation and systems that connect technology with real operations.",
contact:"LET'S WORK",contact1:"Have an idea that",contact2:"deserves to exist?",talk:"Let's talk"
}};

const tech=["Next.js","React","TypeScript","Node.js","PostgreSQL","Supabase","OpenAI","Vercel"];

const featuredProjects = [
  {
    name: "Nexo Ministerial",
    category: "SaaS · Gestión ministerial",
    image: "/projects/nexoministerial.jpg",
    fallback: "/projects/nexo-ministerial.svg",
    href: "https://nexoministerial.vercel.app/",
    className: "project-feature project-feature--wide"
  },
  {
    name: "TecmovilesOS",
    category: "Sistema · Operaciones",
    image: "/projects/tecmovilesos.jpg",
    fallback: "/projects/zzlabs.svg",
    href: "https://github.com/zzlabs-community",
    className: "project-feature"
  },
  {
    name: "Registro de Formularios",
    category: "Aplicación web",
    image: "/projects/registrodeformularios.jpg",
    fallback: "/projects/proyecto-03.svg",
    href: "https://github.com/zzlabs-community",
    className: "project-feature"
  },
  {
    name: "Seguros Cotizador",
    category: "Plataforma · Cotización",
    image: "/projects/seguroscotizador.jpg",
    fallback: "/projects/proyecto-04.svg",
    href: "https://github.com/zzlabs-community",
    className: "project-feature project-feature--wide"
  },
  {
    name: "Adopta Animales",
    category: "Plataforma web",
    image: "/projects/adoptaanimales.jpg",
    fallback: "/projects/proyecto-03.svg",
    href: "https://github.com/zzlabs-community",
    className: "project-feature"
  },
  {
    name: "Abogados",
    category: "Web profesional",
    image: "/projects/abogados.jpg",
    fallback: "/projects/proyecto-04.svg",
    href: "https://github.com/zzlabs-community",
    className: "project-feature"
  },
  {
    name: "Cafetería",
    category: "Experiencia digital",
    image: "/projects/cafeteria.jpg",
    fallback: "/projects/zzlabs.svg",
    href: "https://github.com/zzlabs-community",
    className: "project-feature project-feature--wide"
  }
];

export default function Home(){
 const[lang,setLang]=useState<Lang>("es");
 const[menu,setMenu]=useState(false);
 useEffect(()=>{const q=new URLSearchParams(location.search).get("lang");if(q==="en")setLang("en")},[]);
 const t=C[lang];
 const toggle=()=>{const n=lang==="es"?"en":"es";setLang(n);const u=new URL(location.href);u.searchParams.set("lang",n);history.replaceState(null,"",u)};
 const anchors=["#systems","#services","#projects","#founder","#contact"];
 return <main>
  <header className="nav shell">
   <a href="#" className="brand"><Image src="/logonav01.png" alt="ZZLabs" width={146} height={40} priority style={{ width: "auto", height: "auto" }}/></a>
   <nav className="desktop-nav">{t.nav.map((n,i)=><a key={n} href={anchors[i]}>{n}</a>)}<button className="lang" onClick={toggle}><Globe2 size={14}/>{lang==="es"?"EN":"ES"}</button></nav>
   <div className="mobile-actions"><button className="lang" onClick={toggle}><Globe2 size={14}/>{lang==="es"?"EN":"ES"}</button><button className="menu-btn" onClick={()=>setMenu(!menu)} aria-label="Menu">{menu?<X/>:<Menu/>}</button></div>
   {menu&&<motion.div className="mobile-menu" initial={{opacity:0,y:-8}} animate={{opacity:1,y:0}}>{t.nav.map((n,i)=><a key={n} href={anchors[i]} onClick={()=>setMenu(false)}>{n}<ArrowUpRight size={16}/></a>)}</motion.div>}
  </header>

  <section className="hero">
   <div className="hero-grid"><RippleGrid gridColor="#7c3cff" opacity={.42} glowIntensity={.18}/></div>
   <div className="hero-aurora"/>
   <div className="shell hero-content">
    <ShinyText text={t.heroEye} className="eye"/>
    <motion.h1 initial={{opacity:0,y:50}} animate={{opacity:1,y:0}} transition={{duration:.85}}>{t.hero1}<br/><span>{t.hero2}</span></motion.h1>
    <motion.p className="lead" initial={{opacity:0,y:18}} animate={{opacity:1,y:0}} transition={{delay:.25}}>{t.heroP}</motion.p>
    <div className="actions"><StarBorder href="#systems" color="#7c3cff">{t.explore} ↗</StarBorder><a className="btn" href="https://github.com/zzlabs-community" target="_blank"><Code2 size={17}/> GitHub</a></div>
    <div className="hero-stats">
      {[t.stat1,t.stat2,t.stat3].map((s,i)=><div key={s}><small>{s}</small><strong>{t.stats[i]}</strong></div>)}
    </div>
   </div>
  </section>

  <section id="systems" className="section shell">
   <p className="eye">01 · {t.systems}</p><div className="section-heading"><h2>{t.products}</h2><p>ZZLabs convierte necesidades reales en productos que pueden mantenerse, medirse y evolucionar.</p></div>
   <motion.article className="nexo-card" initial={{opacity:0,y:45}} whileInView={{opacity:1,y:0}} viewport={{once:true}}>
    <div className="nexo-copy"><span className="pill">FLAGSHIP · 2026</span><h3>Nexo<br/>Ministerial</h3><p>{t.nexo}</p><div className="stack">NEXT.JS · TYPESCRIPT · SUPABASE · POSTGRESQL · RLS · VERCEL</div><a className="inline-link" href="https://github.com/zzlabs-community/NEXOMINISTERIAL" target="_blank">{t.view}<ArrowUpRight size={17}/></a></div>
    <div className="nexo-art"><div className="orb o1"/><div className="orb o2"/><div className="nexo-window"><div className="window-bar"><i/><i/><i/><span>NEXO MINISTERIAL</span></div><div className="window-body"><aside><b>NEXO</b><span>Inicio</span><span>Personas</span><span>Asistencia</span><span>Economía</span></aside><div className="dash"><small>RESUMEN GENERAL</small><strong>Una iglesia.<br/>Un sistema.</strong><div className="metrics"><em>Asistencia<b>87%</b></em><em>Ministerios<b>12</b></em><em>Seguimiento<b>24</b></em></div></div></div></div></div>
   </motion.article>
  </section>

  <section id="services" className="section shell">
    <p className="eye">02 · {t.services}</p><h2 className="wide-title">{t.servicesTitle}</h2>
    <div className="service-grid">
      {[[Boxes,"SaaS & Products"],[Workflow,"Automation"],[Sparkles,"AI Integration"]].map(([I,title]:any,i)=><motion.article key={title} initial={{opacity:0,y:30}} whileInView={{opacity:1,y:0}} viewport={{once:true}} transition={{delay:i*.08}}><div className="service-icon"><I/></div><span>0{i+1}</span><h3>{title}</h3><p>{t.serviceDesc[i]}</p></motion.article>)}
    </div>
  </section>

  <section className="focus-section">
   <div className="shell"><p className="eye">03 · {t.focus}</p><TrueFocus sentence={t.focusSentence} borderColor="#7c3cff" glowColor="rgba(124,60,255,.55)"/></div>
  </section>

  <section className="section shell tech-section">
    <p className="eye">04 · {t.tech}</p><div className="section-heading"><h2>{t.techTitle}</h2><p>Elegimos herramientas por estabilidad, velocidad de desarrollo y capacidad de escalar.</p></div>
    <div className="tech-grid">{tech.map((x,i)=><motion.div key={x} initial={{opacity:0,scale:.95}} whileInView={{opacity:1,scale:1}} viewport={{once:true}} transition={{delay:i*.04}}><span>{String(i+1).padStart(2,"0")}</span><strong>{x}</strong></motion.div>)}</div>
  </section>

  <section id="projects" className="projects">
   <div className="shell projects-head">
    <p className="eye">05 · {t.projects}</p>
    <h2>{t.projectsTitle}</h2>
    <p>{t.projectsP}</p>
   </div>

   <div className="shell project-showcase">
    {featuredProjects.map((project,index)=>
      <motion.a
        key={project.name}
        href={project.href}
        target="_blank"
        rel="noreferrer noopener"
        className={project.className}
        initial={{opacity:0,y:34}}
        whileInView={{opacity:1,y:0}}
        viewport={{once:true,amount:.15}}
        transition={{duration:.7,delay:index*.06}}
      >
        <div className="project-feature__media">
          <img src={project.image} alt={project.name} onError={(e)=>{e.currentTarget.onerror=null;e.currentTarget.src=project.fallback}}/>
        </div>
        <div className="project-feature__info">
          <div>
            <span>{project.category}</span>
            <h3>{project.name}</h3>
          </div>
          <ArrowUpRight size={22}/>
        </div>
      </motion.a>
    )}
   </div>

   <div className="shell projects-footer">
    <a className="inline-link" href="https://github.com/zzlabs-community?tab=repositories" target="_blank">
      {t.githubCta}<ArrowUpRight size={17}/>
    </a>
   </div>
  </section>

  <section className="section shell process">
   <p className="eye">06 · {t.process}</p><h2 className="wide-title">{t.processTitle}</h2>
   <div className="process-grid">{t.steps.map(([n,title,desc])=><article key={n}><span>{n}</span><h3>{title}</h3><p>{desc}</p></article>)}</div>
  </section>

  <section id="founder" className="section shell founder">
   <motion.div className="photo" initial={{clipPath:"inset(100% 0 0)"}} whileInView={{clipPath:"inset(0)"}} viewport={{once:true}} transition={{duration:.9}}><Image src="/founder.jpg" alt="Benjamin González" fill sizes="(max-width:800px) 100vw,45vw"/></motion.div>
   <div><p className="eye">07 · {t.founder}</p><h2>Benjamin<br/>González</h2><b>{t.founderRole}</b><p>{t.founderP}</p><div className="founder-tags"><span>Full Stack</span><span>Product</span><span>Automation</span><span>AI</span></div></div>
  </section>

  <section id="contact" className="contact shell"><p className="eye">08 · {t.contact}</p><h2>{t.contact1}<br/><span>{t.contact2}</span></h2><StarBorder href="mailto:contact@zzlabs.site" color="#51d7ff">{t.talk} ↗</StarBorder></section>
  <footer className="shell"><Image src="/logonav01.png" alt="ZZLabs" width={112} height={30} style={{ width: "auto", height: "auto" }}/><span>ZZLabs · Independent software lab · © 2026</span><a href="https://github.com/zzlabs-community">GitHub ↗</a></footer>

  <a
    href="https://wa.me/542923500173?text=Hola%20ZZLabs%2C%20quiero%20consultar%20por%20un%20proyecto."
    target="_blank"
    rel="noopener noreferrer"
    className="whatsapp-float"
    aria-label="Contactar a ZZLabs por WhatsApp"
  >
    <MessageCircle size={21} strokeWidth={1.8}/>
    <span>{lang==="es"?"Hablemos":"Let's talk"}</span>
  </a>
 </main>
}
