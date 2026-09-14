"use client";
import { CSSProperties, useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import "./DriftWall.css";

export type DriftItem = {
  image: string;
  title?: string;
  subtitle?: string;
  description?: string;
  href?: string;
};

type Props = {
  items: DriftItem[];
  columns?: number; tileWidth?: number; tileHeight?: number; gap?: number; radius?: number;
  tilt?: number; turn?: number; roll?: number; perspective?: number; depth?: number; speed?: number;
  direction?: "up"|"down"; variance?: number; parallax?: number; pauseOnHover?: boolean; lift?: number;
  fade?: number; dim?: number; grayscale?: boolean; overlayColor?: string; className?: string; style?: CSSProperties;
};

const prefersReducedMotion = () =>
  typeof window !== "undefined" && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

const columnFactor = (index:number, variance:number) =>
  1 + variance * ((((index * 0.6180339887 + 0.35) % 1) * 2) - 1);

export default function DriftWall({
  items, columns=4, tileWidth=290, tileHeight=190, gap=20, radius=20, tilt=11, turn=-10, roll=0,
  perspective=1300, depth=90, speed=28, direction="up", variance=.35, parallax=.5, pauseOnHover=false,
  lift=50, fade=.45, dim=.8, grayscale=false, overlayColor="#10091d", className="", style
}:Props){
  const containerRef = useRef<HTMLDivElement|null>(null);
  const planeRef = useRef<HTMLDivElement|null>(null);
  const trackRefs = useRef<(HTMLDivElement|null)[]>([]);
  const rafRef = useRef<number|null>(null);
  const offsetsRef = useRef<number[]>([]);
  const velocitiesRef = useRef<number[]>([]);
  const hoveredColRef = useRef(-1);
  const wallHoveredRef = useRef(false);
  const pointerRef = useRef({x:0,y:0});
  const pointerDampedRef = useRef({x:0,y:0});
  const lastTsRef = useRef<number|null>(null);
  const [containerHeight,setContainerHeight] = useState(650);
  const [activeId,setActiveId] = useState<string|null>(null);
  const activeIdRef = useRef<string|null>(null);
  const [reduced,setReduced] = useState(false);

  useEffect(()=>{
    setReduced(prefersReducedMotion());
    const mq=window.matchMedia("(prefers-reduced-motion: reduce)");
    const onChange=(e:MediaQueryListEvent)=>setReduced(e.matches);
    mq.addEventListener("change",onChange);
    return()=>mq.removeEventListener("change",onChange);
  },[]);

  const safeItems = items.length ? items : [{image:"/Logo.png",title:"ZZLabs",subtitle:"GitHub"}];
  const columnItems = useMemo(()=>{
    const cols = Array.from({length:columns},()=>[] as DriftItem[]);
    safeItems.forEach((item,i)=>cols[i%columns].push(item));
    return cols.map(c=>c.length?c:safeItems.slice(0,1));
  },[safeItems,columns]);

  const columnMeta = useMemo(()=>{
    const unit=tileHeight+gap;
    return columnItems.map(col=>{
      const copyHeight=Math.max(unit,col.length*unit);
      const copies=Math.max(2,Math.ceil((containerHeight*1.7)/copyHeight)+1);
      return {copyHeight,copies};
    });
  },[columnItems,tileHeight,gap,containerHeight]);

  useLayoutEffect(()=>{
    if(!containerRef.current)return;
    const ro=new ResizeObserver(([entry])=>setContainerHeight(entry.contentRect.height||650));
    ro.observe(containerRef.current);
    return()=>ro.disconnect();
  },[]);

  const baseVelocities=useMemo(()=>{
    const dir=direction==="up"?1:-1;
    return columnItems.map((_,c)=>speed*columnFactor(c,variance)*dir*(c%2===0?1:-1));
  },[columnItems,speed,direction,variance]);

  useEffect(()=>{
    offsetsRef.current=columnMeta.map((m,c)=>m.copyHeight*((c*.37)%1));
    velocitiesRef.current=columnItems.map(()=>0);
  },[columnMeta,columnItems]);

  const applyPlaneTransform=useCallback((px:number,py:number)=>{
    const p=planeRef.current;if(!p)return;
    p.style.transform=`translate(-50%,-50%) scale(1.12) rotateX(${tilt+py}deg) rotateY(${turn+px}deg) rotateZ(${roll}deg) translateZ(${-depth}px)`;
  },[tilt,turn,roll,depth]);

  useEffect(()=>{
    const animate=(ts:number)=>{
      if(lastTsRef.current===null) lastTsRef.current=ts;
      const dt=Math.min(.05,Math.max(0,ts-lastTsRef.current)/1000);
      lastTsRef.current=ts;
      const maxTilt=parallax*7;
      const targetX=pointerRef.current.x*maxTilt,targetY=-pointerRef.current.y*maxTilt,damp=1-Math.exp(-dt/.12);
      pointerDampedRef.current.x+=(targetX-pointerDampedRef.current.x)*damp;
      pointerDampedRef.current.y+=(targetY-pointerDampedRef.current.y)*damp;
      applyPlaneTransform(pointerDampedRef.current.x,pointerDampedRef.current.y);

      if(!reduced){
        for(let c=0;c<trackRefs.current.length;c++){
          const meta=columnMeta[c];if(!meta)continue;
          const factor=(wallHoveredRef.current&&pauseOnHover)||hoveredColRef.current===c?0:1;
          const target=baseVelocities[c]*factor;
          const ease=1-Math.exp(-dt/(target===0?.16:.28));
          velocitiesRef.current[c]+=(target-velocitiesRef.current[c])*ease;
          let next=(offsetsRef.current[c]??0)+velocitiesRef.current[c]*dt;
          next=((next%meta.copyHeight)+meta.copyHeight)%meta.copyHeight;
          offsetsRef.current[c]=next;
          const el=trackRefs.current[c];
          if(el)el.style.transform=`translate3d(0,${-next}px,0)`;
        }
      }
      rafRef.current=requestAnimationFrame(animate);
    };
    rafRef.current=requestAnimationFrame(animate);
    return()=>{if(rafRef.current)cancelAnimationFrame(rafRef.current);lastTsRef.current=null;};
  },[applyPlaneTransform,baseVelocities,columnMeta,pauseOnHover,parallax,reduced]);

  const activate=(id:string,c:number)=>{activeIdRef.current=id;hoveredColRef.current=c;setActiveId(id)};
  const release=()=>{activeIdRef.current=null;hoveredColRef.current=-1;setActiveId(null)};
  const move=(e:React.PointerEvent)=>{
    const rect=containerRef.current?.getBoundingClientRect(); if(!rect)return;
    if(parallax>0&&!reduced)pointerRef.current={x:(e.clientX-rect.left)/rect.width-.5,y:(e.clientY-rect.top)/rect.height-.5};
    const hit=document.elementFromPoint(e.clientX,e.clientY) as HTMLElement|null;
    const tile=hit?.closest("[data-tile-id]") as HTMLElement|null;
    if(!tile)return;
    const id=tile.dataset.tileId!;
    if(id===activeIdRef.current)return;
    activate(id,Number(tile.dataset.col));
  };

  const cssVars = {
    "--dw-tile-w":`${tileWidth}px`,"--dw-tile-h":`${tileHeight}px`,"--dw-gap":`${gap}px`,
    "--dw-radius":`${radius}px`,"--dw-perspective":`${perspective}px`,"--dw-lift":`${lift}px`,
    "--dw-dim":dim,"--dw-gray":grayscale?1:0,"--dw-overlay":overlayColor,
    "--dw-edge":`${Math.max(0,(1-fade)*100)}%`,...style
  } as CSSProperties;

  const tile = (item:DriftItem,id:string,c:number) => {
    const inner = <span className="drift-wall__inner">
      <span className="drift-wall__fallback"/>
      <img
        src={item.image}
        alt={item.title??""}
        loading="lazy"
        decoding="async"
        onError={(e)=>{(e.currentTarget as HTMLImageElement).style.display="none"}}
      />
      <span className="drift-wall__overlay"/>
      <span className="drift-wall__meta">
        <strong>{item.title || "ZZLabs"}</strong>
        <small>{item.subtitle || "GitHub project"}</small>
      </span>
    </span>;
    const common={className:`drift-wall__tile${activeId===id?" is-active":""}`,"data-tile-id":id,"data-col":c,onFocus:()=>activate(id,c),onBlur:release};
    return item.href
      ? <a key={id} href={item.href} target="_blank" rel="noreferrer noopener" {...common}>{inner}</a>
      : <div key={id} tabIndex={0} {...common}>{inner}</div>;
  };

  return <div ref={containerRef} className={`drift-wall ${reduced?"drift-wall--reduced":""} ${className}`} style={cssVars}
    onPointerMove={move} onPointerEnter={()=>wallHoveredRef.current=true}
    onPointerLeave={()=>{wallHoveredRef.current=false;pointerRef.current={x:0,y:0};release()}}>
    <div ref={planeRef} className="drift-wall__plane">
      {columnItems.map((col,c)=>{const meta=columnMeta[c];return <div className="drift-wall__col" key={c}>
        <div className="drift-wall__track" ref={el=>{trackRefs.current[c]=el}}>
          {Array.from({length:meta.copies}).flatMap((_,copy)=>col.map((item,i)=>tile(item,`${c}-${copy}-${i}`,c)))}
        </div>
      </div>})}
    </div>
  </div>;
}
