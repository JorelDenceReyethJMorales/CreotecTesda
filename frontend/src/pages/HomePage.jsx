import React, { useState, useRef, useEffect } from "react";
import Sidebar from "../components/Sidebar";
import { FaChevronLeft, FaChevronRight } from "react-icons/fa";
import { useNavigate } from "react-router-dom";

function ScrollArrow({ direction = "right", onClick, visible }) {
  const [hover, setHover] = useState(false);

  const baseStyle = {
    background: "transparent",
    border: "none",
    color: "#a361ef",
    cursor: "pointer",
    fontSize: "2.2rem",
    userSelect: "none",
    padding: "0 0.5rem",
    transition: "color 0.3s ease, transform 0.2s ease, opacity 0.3s ease",
    outline: "none",
    position: "absolute",
    top: "50%",
    transform: hover
      ? "translateY(-50%) scale(1.3)"
      : "translateY(-50%) scale(1)",
    opacity: visible ? (hover ? 1 : 0.7) : 0,
    zIndex: 10,
    pointerEvents: visible ? "auto" : "none",
  };

  baseStyle[direction === "left" ? "left" : "right"] = 10;

  return (
    <button
      onClick={onClick}
      style={baseStyle}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      aria-label={direction === "left" ? "Scroll left" : "Scroll right"}
    >
      {direction === "left" ? <FaChevronLeft /> : <FaChevronRight />}
    </button>
  );
}

function HomePage() {
  const certScrollRef = useRef(null);
  const tesdaScrollRef = useRef(null);

  const [showCertLeft, setShowCertLeft] = useState(false);
  const [showCertRight, setShowCertRight] = useState(false);
  const [showTesdaLeft, setShowTesdaLeft] = useState(false);
  const [showTesdaRight, setShowTesdaRight] = useState(false);

  const [hoverCert, setHoverCert] = useState(false);
  const [hoverTesda, setHoverTesda] = useState(false);

  const certificates = [
    "Certificate 1", "Certificate 2", "Certificate 3",
    "Certificate 4", "Certificate 5", "Certificate 6", "Certificate 7",
    "Certificate 1", "Certificate 2", "Certificate 3",
    "Certificate 4", "Certificate 5", "Certificate 6",
  ];

  const tesdaRecords = [
    "TESDA Record 1", "TESDA Record 2", "TESDA Record 3",
    "TESDA Record 4", "TESDA Record 5"
  ];

  useEffect(() => {
    const el = certScrollRef.current;
    if (!el) return;

    const handleScroll = () => {
      setShowCertLeft(el.scrollLeft > 0);
      setShowCertRight(el.scrollLeft + el.clientWidth < el.scrollWidth);
    };

    handleScroll();
    el.addEventListener("scroll", handleScroll);
    return () => el.removeEventListener("scroll", handleScroll);
  }, [certificates]);

  useEffect(() => {
    const el = tesdaScrollRef.current;
    if (!el) return;

    const handleScroll = () => {
      setShowTesdaLeft(el.scrollLeft > 0);
      setShowTesdaRight(el.scrollLeft + el.clientWidth < el.scrollWidth);
    };

    handleScroll();
    el.addEventListener("scroll", handleScroll);
    return () => el.removeEventListener("scroll", handleScroll);
  }, [tesdaRecords]);

  const scroll = (ref, direction) => {
    if (!ref.current) return;
    const scrollAmount = 200 * direction;
    ref.current.scrollBy({ left: scrollAmount, behavior: "smooth" });
  };

  const navigate = useNavigate();

  const logout = () => {
    localStorage.removeItem("token");
    // Add any other logout logic here if needed
  };

  const containerStyle = {
    backgroundColor: "#696b6c",
    color: "white",
    margin: "2rem auto",
    padding: "1rem",
    width: "calc(100% - 256px)",
    marginLeft: "256px",
    boxShadow: "0 5px 7px #0000004d",
  };

  const scrollWrapperStyle = {
    position: "relative",
    display: "flex",
    alignItems: "center",
  };

  const scrollContainerStyle = {
    display: "flex",
    overflowX: "auto",
    scrollBehavior: "smooth",
    scrollbarWidth: "none",
    msOverflowStyle: "none",
    flexGrow: 1,
    padding: "0 2rem",
    margin: "0 1rem",
  };

  const itemStyle = {
    flex: "0 0 auto",
    backgroundColor: "#4c3a91",
    margin: "0 0.5rem",
    padding: "1rem",
    borderRadius: "4px",
    minWidth: "150px",
    height: "58px",
    textAlign: "center",
    boxShadow: "0 3px 5px #00000066",
    opacity: 0.85,
    transition: "transform 0.2s ease, opacity 0.2s ease",
    cursor: "pointer",
  };

  const sectionTitleStyle = {
    fontSize: "1.8rem",
    fontWeight: "700",
    marginBottom: "1rem",
    color: "#f3eaff",
    textShadow: "1px 1px 3px rgba(0, 0, 0, 0.5)",
    letterSpacing: "0.5px",
    borderBottom: "2px solid #a361ef",
    paddingBottom: "0.5rem",
  };

  return (
    <div className="flex min-h-screen bg-[#1f1f1f] text-white">
      {/* Sidebar on the left */}
      <Sidebar />

      {/* Main Content on the right */}
      <div className="flex-1 p-6 overflow-y-auto">
        {/* Logout */}
        <div className="flex justify-end mb-4">
          <button
            className="rounded-md bg-red-600 text-white px-3 py-1 hover:bg-red-700"
            onClick={() => {
              logout();
              navigate("/login");
            }}
          >
            Logout
          </button>
        </div>

        <div style={containerStyle}>
          <h2 style={sectionTitleStyle}>Recently Made Certificates</h2>
          <div
            style={scrollWrapperStyle}
            onMouseEnter={() => setHoverCert(true)}
            onMouseLeave={() => setHoverCert(false)}
          >
            <ScrollArrow
              direction="left"
              onClick={() => scroll(certScrollRef, -1)}
              visible={hoverCert && showCertLeft}
            />

            <div ref={certScrollRef} style={scrollContainerStyle} className="scroll-container">
              {certificates.map((cert, i) => (
                <div key={i} style={itemStyle} className="scroll-item">
                  {cert}
                </div>
              ))}
            </div>

            <ScrollArrow
              direction="right"
              onClick={() => scroll(certScrollRef, 1)}
              visible={hoverCert && showCertRight}
            />
          </div>
        </div>

        <div style={containerStyle}>
          <h2 style={sectionTitleStyle}>TESDA Records</h2>
          <div
            style={scrollWrapperStyle}
            onMouseEnter={() => setHoverTesda(true)}
            onMouseLeave={() => setHoverTesda(false)}
          >
            <ScrollArrow
              direction="left"
              onClick={() => scroll(tesdaScrollRef, -1)}
              visible={hoverTesda && showTesdaLeft}
            />

            <div ref={tesdaScrollRef} style={scrollContainerStyle} className="scroll-container">
              {tesdaRecords.map((record, i) => (
                <div key={i} style={itemStyle} className="scroll-item">
                  {record}
                </div>
              ))}
            </div>

            <ScrollArrow
              direction="right"
              onClick={() => scroll(tesdaScrollRef, 1)}
              visible={hoverTesda && showTesdaRight}
            />
          </div>
        </div>

        <style>{`
          .scroll-container::-webkit-scrollbar {
            display: none;
          }

          .scroll-item {
            transition: transform 0.2s ease, opacity 0.2s ease, height 0.2s ease;
          }

          .scroll-item:hover {
            transform: scaleX(1.05);
            height: 70px;
            opacity: 1;
          }
        `}</style>
      </div>
    </div>
  );
}

export default HomePage;