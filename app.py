"""
术语表辅助翻译工具 — Streamlit Web App

Upload a document (.docx / .pdf) and a bilingual glossary (.xlsx),
get an AI translation that strictly follows your glossary.

Translation is powered by the DeepSeek API.
Your glossary stays on the server — users only see the translated output.
"""

import streamlit as st
import tempfile
import os
import base64
import zlib
import json
from pathlib import Path
from openai import OpenAI
from docx import Document
import pdfplumber

# Import the glossary helper (must be in the same directory)
from glossary_helper import load_glossary, protect_terms, restore_terms

# ---------------------------------------------------------------------------
# Embedded glossary (compressed & base64-encoded — not readable by humans)
# Generated from glossary.xlsx at build time.
# ---------------------------------------------------------------------------
_EMBEDDED_GLOSSARY_B64 = """eJy1W1lzFFeW/isZPEHExRjc4Wb8hmHcjcN2KKCfZmJiIpWVpUqoyqzJRVju6AgJECrtkiWQQEILSIi1hCwEUmn7L+26mVlP/gt9zj03l8qqEnhmHOFFlXm3c+539pN/P9VTtBxHtftOfaX8/VS9NsqrS/XdiWDqXjC7Bc9OfaNqrqOoZk75i215Zs45xZRT/vZ9fjRa3x0JHt3llRe8us0Ht4PtfZxQ1HvUomKYrm7rjouj0yPqtRoOil7rOaWs2m6fGDa5i+tWl8IP73DMDS9naAasVbbhDIami52X1/35CX+1Ut9/T6fFoaqj6barGqbiFnQlL45s5cUPTXXERBpM03EKDjqbN8ycYfZEJFUf1fefhqvP5SG/E5Rc08twUt10VdfopTNMVPzHd5qGXvLcgmUbPwE9bca/WQuHdnDYNd0p65qb94rFPsXxukuGiyxwLUH/4Ad/cZhYm2YZ8Bjo5WsP5F6qqVie6xg53VZuFSyloOYUVWnDduDkzvNw4ym/tx3UVoLaB5wNuzquauLZ4Aww1ir2AgsEr3KGU/Zccebwyatwc7h+9A4prW3wmQGxdRG2gQMbZt6GSTivbFs5T3NPO2fEnhuz/vA8zAyrB8AggFHEINe1bFPvE/DhU5WwWuW1N/hGs3BRQ3V1gIKpF+noq4jA/UqwP0NjTK3o5fRcwqfG0Dhf2+CHW3xmnI8sECLEWLhwvHzH6DGNvKEBrXjeoqebBKGwuoY3TXjYHfTXH4lpRV21gSPAPLW7aDgFvJcITDit0X87mH1O1+O/eRqsVulktg03ChPVcrloyElF9ZbA3OGkP9zPF98KDuRywG0Sn9WKlK7LqkenvQQIt0xxvOk5fni/sboXLIgdug04jtVjq+WCoSmGq5ecBM+Slt1xIAQHe6agAwiGkY4i8S2EAQlpoWNxONzcJIGUlERTBTmaiqeKRElSFTwf5bVJuENYQwi8egsAHc0DPGh6zgNSxdj9mbBaa2FXNJju1YE9CBeDr4Pldb67G27eo9GlEmx/07RuFfVcj65IAQeNITi1OeCPV2FlPnQcibQHQhBR/dlnn8FQUAGCJeeUGPoqCl033HUvCJFERfD6GZEE/CQK5QnwVh3lluEWiAlCzjJUHtTqe5N0I/znMakdjx7D6YhB2k0lOhxqUlqjW3UMInvhkFffA+VANsgaSBzf6hfoKBimqvwgLgEmXAXhBlmhdbpsqww6r0+5lCsZpuG4Nt3V6cs/XO26JITRX9nlx3fqu/v8YEUyH7BWhtHAb+CH+K0VVaMVUTEVhJbG7EOJEyQjgRUSIy6kGSyACCVifgwscaKB7XBTqKECqKlu1D//44HEif3Dnaf1/TV8admg26QKazwcDAdmpW2RPATYmLoQGKCCnqEAOo6libsVU8cnQccE1dXw6IH/eNV/fSQW6LF1uPsbnm04YF3EEsABz5GK/ef6/gO/f4Om4QT9RwHQXr1pjpCjx0ttxpuWebbzHD50DzAh5UYyrbsYS5bU/fd+CV4OhBt3w7E7UvnlYSCZNGQtgK/XiIUGhJ9Pbvo/rzVmI1VjAMN7LVy6L1GYQe2Z/3gZVCUpFXGDip7Py6vhlcf+m3USKRxRto2Siqg19CY56YjyEwAeDr0M7r6XlsBUXc9GEdJth7jiVz4EW+NCUz6XFJRBZvRI79DIf/Y/dJSCBVdVsIo5QE5PDHrSNJOgOPjWIChduUhB7TbEkUBMEwLCnXeNh0/wPa592fJImfiLr0Etx4+l1wCkorJ88JYvgbszA2YZpFMoWtcoGiB9JSunAy2AOjAzttFTcNtP6zRHbP1+GM1XPK5LvDmrumcNx/GI68IQk6IWplcvRZNRWwgnyLplAqMKRllMOH7sj6wBgviC8DwcDdQFMhTA4+qJsZk65IPr/+x/Bv9kp+RBqoTek1q3zRLwC+RcHhAcFn+k31/cDDbmOWiyvVHwAwRDbTUH03XN1l0lfX6U3eggVVD6NYTwrECTBcDEcUreshVHLdJtDG7xqbGEFzcsw2xeUXBkbTwYfoHMnxkINsYF5nSQmiJ6Ar2qbeiuQXq7sXDPr0wFs+/CD5t8F+g+Cqvwz2rkoYLRRUuhGbbmGS7IXh84XkpOR9eCNPfSXf/Bm8S5AmtllS3QRdKTcqIbqryQ6BcQQRwXpYrzx5fR05ua8JcWcUBOLak9BNTjBT75lN8XptAolS3bpSnEJfnCQZNAynJyGnzbGKX+3Ip/vxKujvmzx8kJ09YCFBRdAAhI2v/DOzbMXmAnvpPYhgdq0cjF8kbL8+pTPj0COt3fuk1cS0ZFyr39aHDxwBkH1xSsDc68mp55RdeEgiPB3vHHBkA8JWXgIt8XcpoWOsUAnupFdKWVbl0RSyW8D58NwI71vWfCqMYGL9zcQYbR4yu6o9lGOSZveh5mCTXdjdySvtPDuwAScU22egtNYHaTeAdx9NHXwavR7GtQvXpZh//A4bVoaH1/AsQ4O7TdQMEFPvkhxQLTK3UTBoJZNLD+3LoQYAO8JkUYSfgThQnYm4pgwN2C65CDAeumGw3HcULSve7Irou9D+b5YDTzoN9/MATXIs6RjFNc/Uc31ty2AVGRDO8+MkeM2dzzN2dhDPjy58+lf11o+vUFaSxxb+fPXblw7soXUg/5i/3+xHp4PB1MvEWoCIVE/r3lABjpSJEHycdqvPIqGN7jR+KuQQBccFiEaOhFoVHQUUo/1goqogGUEzzTnJSX8mxAxloQr4ltG3PVxpN5oQF1rWACscJxBINfkjiNfReMwshvAb3VeII+Bc77wYIoTbKvstDoX5HPr5J0wn9kRAEOAx+qYTQugj50SYb3gtt7ZMktkHnESSr2U/I62mGpoeaPIOqA2aBywAkgP8YVnglM7onDFrDQW0uNJ/f48jJRCEqNTxOFBWCv4tw0ikURH0Z+OnCPzHebd6T334aHh7h1q6kCtuOtwSkFiQs7cLO/9g+AzKLPv/jWn3jONw+Abrq6KA4D/8MzQfUY6AIhcyFu9GT2QIR7sAavjvn3d3BboQMoMM5DtGgAn2AKhQy4lIgAdBkU3znka8to3CeEg6Gbgkfo6sJ9aqpt92FoDjTbOmign6QneozYkKYxhhrwFK8QIJGAI30n9f31WD/Vd2uJnwWwcbyysAYiiBGKO6O8xF7+g73Io4l3AJx5mVHDo371XfOokq5GAdkgr9wjyW3hkbweIrLxaIpPvUpGGhg3aUbLOJK4LDPSApYcQ0ogifVwsDASSxZYAGkuu0EE8pQmSk0UXi0JxiLelxQb1QNmAdO1ZpU0t+4vLpPtbav2/P118Ob47m2+WPNnNxuD46ACE5N6JbKavbqtXOqWAf33qn0T3J0rKHuqdD3bLkQ2Rs4q0axcelbj4ZTIV6z58yIZ4vahKgewqdLpicMAIGdvu/Fqvr7/gZIBedXAnEYBTCQqPk3mushtAdcAgJVyW4RGw+AKEzrJOP/tNDj9bZwczeg1ik2uDrlmycKuJTQA/s8Az12RZpSGBfsz8hLj94B9r0h64fiQj6wEr8HqD4OWIbsmTXUXOHBkNiyIfjH59beCYeeUriiTKEeTxxWNJgMZPnnV9BYzcpYwsYLgiWWwiGIW2FtASj6fPK7v9p+Tf9XGor92hxNTFM9Rrp5L/d3846o0VHAQudUVPQ9GXqUT0GPcSv4FW0XPUlvFc2Cr1N/NP6KtYi6SokZWxUlXcMzISyGvDK+lw2jltMiBmJbS4sCcIW8Kk37+yhBEcTLReANTPSXVBWuJ4FbNlMPPJ2/DBdMkcmVsFEz0aGITEYX4bd5BIGX0JMGfvNLRxuoe3WrieQr08soaWObIAxS4/dbL9cTqRQwKnw6goKUHXfOKUX745zW+sBS++yCDGmQGOoGgbpL0mANhgiCu9gb8W8y+bg5QCA9RpqLZluOc1X9UUbijXEk6tuaHM3AMYH3jYSU8fhuxBdwDF0hPifoRrFvf3a8fzqQ5DeKg5tDUZRJbcJ/9nKKavGE76MujGyB1ksh4gYGS6VPcygOBpNwMGNV420b/I5BsPvBI5i9kdq4pOz987I8OSm5iOoeyNhDFluA6zUhlFA2IzItSCUvcUJC0N8xHVqXTIhJsGL1TyITGGHwneetRvhAZQRf2A3gsNtqBKzqZmvruCKawZf65XIZAKRGyNi91PZJ1uLpW6YyQ941g4dUUC2lKG93ReQqla0RAqPdaN+miNvYb/QMyYDadvPTo49RSSzKpck/UBe4GUxQJgsXFUwIMHB1wg9KCyXBi1XTFH1mXaCkRSv2xIUx2xhtQYIvTpEdTUNvlsEQQngguBeGJYEcptzv84HZYrWZPjvIi4Brnr1I2oyn7cwtzhIJ/qXIDph8NTByXqHaS0pnpZHsyN5fcnyPsqoZRkcDzi22QaAKCzMeB1+aIrCUYtZLhSJ+FIlC4X1IQrXKk2EJN4A4qOEh2CTwzTDc7WJKK1qggrFILICfsWL0ggCprjfsikd0lEla4IiopPXkvne34qRDeeT7yQkJAvFAuOY4h/HzB3r1n4GPLiZeLun2TZKO/6TiSACwTtGgIrAgezlBJRgIojiJclDhxkskp0I1ZZsa1LUpDNGnhTN46rY7pQsOjF8HG22YZVO0eTzL0zSwMkI4bhBiGlRMxAw6VubvJp3xtJzm3hcqkAH59xHIgqgZe0PNgrRZpPopPtCgpSBAhrCdJDkqZxxnuJqxQHYpPzsGm4Y4wRqIglbM0cOxEQkuXedQa6MnGw5d4s3Q7XboFoYbIc37r2ZadKLFg5H3CVKXbNvS85BO8THIoUdqkOXUYaTo0SlMTVJXz706GGwNNilHs22U57tm/Eo+Ur6N9SME1OQYpj6sDakgpihsn7hs9AvM3EKPxTQtcNBOZyVAJmpMVAfIyfSv2PBvLYHxpGQXU5Dp+26yJLhewoglutLAu31uRV3mFrjQFAWH3m5ZqOzrtXv52UIl9uN8Ohpt9UeV02lM7Iy/Kf/chXQZKaL1eAGe/oBodakBxFhvTTWATBPwEMzVNL0s1ma2FdCiCYIJpa+lTKiDpEovYrLW4kuwbHi9B+C5TElEBB8UBNLqF1bmMkUhqPOR73PZXDhKrXBaUJ8qdDo4OUfjuLoyU8hxX9+kMBIjG/HsZ8TZveVaL0QBTpdpdfMuXlxMSwZO1DTivIYDcTCdcOZqwyvtG/0obmtGHlZHqiaTygxpKxehw4/FKsnHessEfAphq2GdAqrqlIiWgSn4IpjZiX1VqKCV2TISsi8wGPpB5MwwchVIi2PHB7Xptgup/v/ZD+PGGj0/QAHQBfh7jC4fS0yMPbvv+r/2juO/G08bQCKWpgF9lW3ebysjXvTJ6lUpK3Qn8YuoR319qrTyTM94VFV2V79Rb0avUKtd0CtzxlSidnkQY2YXWCJcUeLpIHo6u871fKJn3fyTvut6roxW6isoZn4Piss2o/SJDOBIpHci/CcAhWYIRWJ1ylEtw+0KCbKsEuksG5TTiSrr+IMrAlDhJ6nui+JtKmTSV+LDfg0ohfOEoqO1KkcO8gCx/KOkEuKygTCwjRyk/HtVMHK2gQwQoM6XhyGYwPMRHlsM7h0nmTSsYZSXvmaSJUKHbXuIfxhlNmfrZnOWbey1+SEwIohrMlywtwmUFK5Ri0jUr0gbw1F+akuGVE7kDY3uNQVEtgvNqBdBI4v/R27UX/tz7+tECr802Zo4CgR7Q8P6HLb64DeodW1IeTmEKZlKw6/L31/EWwEHq9qgwf72gmz/Bv2eU7wxwXmU+bA+s32KwOEAL8d3bweGcv/eAj9X81dt8AfyTucbATNg/wBd2grUjeB7UnvDqY2D4+S/+jU9+aAxNN44e0sbn//T5eTjU+T9duMira9LoXLMAH/jmND4/w5QfrM8UmKv81TJ7ulVLuWapOQZws27pBmCyVPJMiM2Y8hfP6PMgaL/udecEJZoLwzyr4EnK8HdEFUtELhzd85cWERR3p8OdXX9+Bs5E1q0NoxD1WIctIWBtgLoqPN/Tkck7A+8+g33dHBUaQQ8tHKYXxOry7BZtFG7eb7wcAyY1XrzmU8PApIsXgUeX4cmX8H9//YAPPud7T347GLvw+efnvziPJ/jaM4pizy+Z8nURHDXlMjHp4kXle/WGIRnUZRQtV/kGbJw4pK78BxhyJTomS/1FSycM4YtPw+OXfHA02JgGFJ8EG2RnT864CXsgpK2i1dOXhk4TL1rBg0j4ZQ6TrI9fgu0i8CDLju8QePjmGPwNfLnwZ+DHpS/5xG2/cvzF5+djpMDfwIaIJZe+JFZc+LPyjfcjqCPixWXw31CRpdDSZeDBM2ABhPUAV07GS3A0AVcaTj2v7x7y6jB/vpjiB8b0pqp8993lJmjxN1Okpvh98LN3gaV8bK5ee90RYP+OTpNtgZJw0JCgKnPE8XUwrMrpr3XjBhCQZvA/YD9HpNn/2wEFroE2/Ur5zz9cYv8AjP9/Q+UPNKt/iCvCfn9l8X+Dtd9tJn6frvzERPQnJpF/h1pin9ibyj6lh4x9SqMC+2jHAPto2TU+dftOUXZy2x47ueD4karQRxLrH4nKPxJHfiTUYCe0Z7MTOn7YCU087IReGXZC9ZmdUDRmJ5Rj2QkFoJMi/JOi6JMMDevUucw6dWeyTp1wrFPTGOvUysI6NVawTu0MrFN5mHWqlHbCa+cMR8d4sqNLztp258fDmxrR4+RVU4c5a9uczdr2XmeUjWyBzo6l/mbWtr03k8GQiRHWtrMzFqqm5sy2d91eeYkztLYDsbadP6xt0wBrWz1nbcu0rG0thrWttrTPc7ZPmrTPT7QPMttHku2jwPg2k29PWOuHI6z1Mw3W+jkGa23kZa1tsrF5SPpiWWsrKmttLGWtHYastdEsPlfSZMVaG21YawMMay3ls9YqZvwoqVmy1vphq6qUWr6p4MJaKyistejBWpP4rDXdzVoT26w1f8xaE79t8oRtEgss85VQDBv5XQ/LfDrCMp9TsMzXCyzzhQLLfIPAMh8ZxECXTfWJF0It6izTR80yzb0s05jLMo2vLNPjGgt3SqzT/aIZ4MnzN1mQphY3lulhYpleoeQ3dQWxTAsQyzT6JPtRDw3LdKvE55ctIRn0Jb9lL0jG80kpLGorYJm+AZYp/bNMnZ5lSvEsUzyPFYqsf7NMlZtlasUxlmUmOv4ti78sU5VlmSJrfL/J+ZuqkCxTKIzOF5X4WKZyxzLVOZYpxLFMLY1lqmIsUwVjmZpSQh/VceLfKZ8qXXFgmSoDy5QTMiZXmusks88yWXyWycCz9LdjLP2BFUt/XMHSnxzEMhIBKm6rZulWapbuB2bpPtsYZJESjzs0Wbotk6WbzVLArJH7kdJ9SUcGS/cwsHTLAkuXhOMfUp8mNV2WLoTGC0iFKr87ZPEHZyz+AofFn9uw+IsIFn8EEYtlxLwt4cfJnmgW90GzuGMqBpoUQWxuYHFLQwwLMYM6EVicnGVxQpbFWdj/YqJpxnRPfaVcOP/5P/4FOdVb8w=="""


def _decode_embedded_glossary():
    """Decode the embedded glossary from compressed base64."""
    compressed = base64.b64decode(_EMBEDDED_GLOSSARY_B64)
    json_str = zlib.decompress(compressed).decode("utf-8")
    return json.loads(json_str)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="TransLegal",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
/* ===== Global ===== */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', 'PingFang SC', 'Hiragino Sans GB', 'Microsoft YaHei', sans-serif;
    color: #2D2A26;
}

/* ===== Hero Banner ===== */
.hero-banner {
    background: #FFFFFF;
    padding: 2rem 2.5rem;
    border-radius: 12px;
    margin-bottom: 1.5rem;
    text-align: center;
    border: 1px solid #EBE7E1;
    box-shadow: 0 1px 2px rgba(0,0,0,0.04);
}
.hero-banner h1 {
    color: #2D2A26 !important;
    font-size: 2.2rem;
    font-weight: 600;
    margin: 0;
    padding: 0;
    letter-spacing: -0.02em;
}
.hero-banner p {
    color: #6B625A !important;
    font-size: 1.05rem;
    margin: 0.4rem 0 0 0;
}
.steps-row {
    display: flex;
    justify-content: center;
    gap: 1.5rem;
    margin-top: 1.25rem;
}
.step-badge {
    background: #FAF8F5;
    border: 1px solid #EBE7E1;
    border-radius: 8px;
    padding: 0.5rem 1.25rem;
    color: #6B625A;
    font-weight: 500;
    font-size: 0.9rem;
}
.step-badge .step-num {
    display: inline-block;
    width: 24px;
    height: 24px;
    line-height: 24px;
    border-radius: 50%;
    background: #C77D4F;
    color: #FFFFFF;
    text-align: center;
    margin-right: 6px;
    font-weight: 600;
    font-size: 0.8rem;
}

/* ===== Cards ===== */
.settings-card {
    background: #FFFFFF;
    border: 1px solid #EBE7E1;
    border-radius: 12px;
    padding: 1.5rem;
    box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    margin-bottom: 1rem;
}
.settings-card h3 {
    margin-top: 0;
    font-size: 1.05rem;
    color: #2D2A26;
    border-bottom: 1px solid #EBE7E1;
    padding-bottom: 0.75rem;
    margin-bottom: 1rem;
    font-weight: 600;
}

/* ===== File Uploader Override ===== */
[data-testid="stFileUploader"] {
    border: 2px dashed #D9D2C7 !important;
    border-radius: 10px !important;
    padding: 1.25rem !important;
    transition: border-color 0.2s, background 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #C77D4F !important;
    background: #FDFCFB;
}

/* ===== Buttons ===== */
.stButton > button {
    border-radius: 8px !important;
    font-weight: 500 !important;
    transition: all 0.2s ease !important;
    padding: 0.5rem 1.25rem !important;
}
.stButton > button:hover {
    transform: translateY(-1px);
    box-shadow: 0 2px 8px rgba(199, 125, 79, 0.25);
}

/* Primary button */
.stButton > button[kind="primary"] {
    background: #C77D4F !important;
    border: 1px solid #B86D3F !important;
    color: #FFFFFF !important;
    font-size: 1rem !important;
}
.stButton > button[kind="primary"]:hover {
    background: #B86D3F !important;
}

/* Download button */
.stDownloadButton > button {
    border-radius: 8px !important;
    font-weight: 500 !important;
    background: #C77D4F !important;
    border: 1px solid #B86D3F !important;
    color: #FFFFFF !important;
    transition: all 0.2s ease !important;
}

/* ===== Metrics ===== */
[data-testid="stMetric"] {
    background: #FAF8F5;
    border-radius: 8px;
    padding: 0.75rem;
    border: 1px solid #EBE7E1;
}
[data-testid="stMetricValue"] {
    color: #C77D4F !important;
}

/* ===== Text areas ===== */
.stTextArea textarea {
    border-radius: 8px !important;
    border: 1px solid #EBE7E1 !important;
    background: #FDFCFB;
}

/* ===== Select / Radio / Input ===== */
.stSelectbox [data-baseweb="select"] > div,
.stTextInput input {
    border-radius: 8px !important;
    border-color: #EBE7E1 !important;
}
.stRadio [data-testid="stMarkdownContainer"] p {
    font-weight: 500;
}

/* ===== Progress bar ===== */
.stProgress > div > div {
    background: #C77D4F;
    border-radius: 4px;
}

/* ===== Expander ===== */
[data-testid="stExpander"] {
    border-radius: 8px !important;
    border: 1px solid #EBE7E1 !important;
}

/* ===== Info callout ===== */
div[data-testid="stNotification"] {
    border-radius: 8px !important;
}

/* ===== Hide sidebar default ===== */
[data-testid="stSidebar"] {
    display: none;
}

/* ===== Footer ===== */
.app-footer {
    text-align: center;
    color: #A39688;
    font-size: 0.85rem;
    padding: 2rem 0 1rem 0;
    border-top: 1px solid #EBE7E1;
    margin-top: 3rem;
}

/* ===== Tab styling ===== */
.stTabs [data-baseweb="tab-list"] {
    gap: 0.25rem;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0;
    padding: 0.5rem 1rem;
    font-weight: 500;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Hero Section
# ---------------------------------------------------------------------------
st.markdown("""
<div class="hero-banner">
    <h1>⚖️ TransLegal</h1>
    <p>上传文档 → AI 自动匹配术语并翻译 → 下载术语一致的译文</p>
    <div class="steps-row">
        <div class="step-badge"><span class="step-num">1</span> 上传文档 & 术语表</div>
        <div class="step-badge"><span class="step-num">2</span> 选择翻译选项</div>
        <div class="step-badge"><span class="step-num">3</span> 下载译文</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Main Setup Area — two columns
# ---------------------------------------------------------------------------
col_left, col_right = st.columns([5, 4], gap="large")

with col_right:
    st.markdown('<div class="settings-card">', unsafe_allow_html=True)
    st.markdown("### ⚙️ 翻译设置")

    _server_key = os.environ.get("DEEPSEEK_API_KEY", "")
    api_key_input = st.text_input(
        "🔑 DeepSeek API Key",
        type="password",
        placeholder="已预填，无需修改" if _server_key else "sk-...",
        help="已从服务器环境变量自动加载。也可填入自己的 Key 覆盖。",
    )
    effective_key = api_key_input.strip() if api_key_input.strip() else _server_key

    direction_label = st.radio(
        "🌍 翻译方向",
        options=["自动检测", "中 → 英", "英 → 中"],
        index=0,
        horizontal=True,
    )
    direction_map = {"自动检测": "auto", "中 → 英": "cn2en", "英 → 中": "en2cn"}
    direction = direction_map[direction_label]

    model = st.selectbox(
        "🤖 翻译模型",
        options=["deepseek-v4-flash", "deepseek-v4-pro"],
        index=0,
        format_func=lambda m: {
            "deepseek-v4-flash": "⚡ V4 Flash — 快速便宜，适合日常翻译",
            "deepseek-v4-pro": "💎 V4 Pro — 旗舰质量，适合正式文档",
        }.get(m, m),
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # Hidden tip
    st.caption("💡 术语表文件不上传时，内置术语表（210条）将自动使用。")

with col_left:
    st.markdown('<div class="settings-card">', unsafe_allow_html=True)
    st.markdown("### 📂 上传文件")

    doc_file = st.file_uploader(
        "📄 文档（.docx / .pdf）",
        type=["docx", "pdf"],
        help="上传待翻译的文档",
    )

    glossary_file = st.file_uploader(
        "📖 术语表（可选 .xlsx）",
        type=["xlsx"],
        help="格式：第一列 = 源语言术语，第二列 = 目标语言翻译。不上传则自动使用内置术语表。",
    )

    # Show file info after upload
    if doc_file:
        file_size_kb = len(doc_file.getvalue()) / 1024
        st.caption(f"✅ 已选择：**{doc_file.name}** ({file_size_kb:.1f} KB)")

    if glossary_file:
        st.caption(f"📋 自定义术语表：**{glossary_file.name}**")

    st.markdown('</div>', unsafe_allow_html=True)

    # --- Translate button ---
    can_translate = doc_file is not None and effective_key.strip() != ""

    if not can_translate:
        if doc_file is None:
            st.info("👆 请上传待翻译的文档开始")
        elif not effective_key.strip():
            st.warning("👉 请在右侧填入 DeepSeek API Key")

    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        translate_clicked = st.button(
            "🚀 开始翻译",
            type="primary",
            disabled=not can_translate,
            use_container_width=True,
        )

# ---------------------------------------------------------------------------
# Translation logic
# ---------------------------------------------------------------------------
def extract_text(file_bytes: bytes, filename: str) -> str:
    """Extract text from .docx or .pdf file bytes."""
    suffix = Path(filename).suffix.lower()
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    try:
        if suffix == ".docx":
            doc = Document(tmp_path)
            text = "\n".join([p.text for p in doc.paragraphs])
        elif suffix == ".pdf":
            with pdfplumber.open(tmp_path) as pdf:
                text = "\n".join([page.extract_text() or "" for page in pdf.pages])
        else:
            raise ValueError(f"不支持的文件格式: {suffix}")
    finally:
        os.unlink(tmp_path)

    return text


def chunk_text(text: str, max_chars: int = 6000) -> list[str]:
    """Split text into chunks at paragraph boundaries, each ≤ max_chars."""
    paragraphs = text.split("\n")
    chunks = []
    current = ""

    for para in paragraphs:
        if len(current) + len(para) + 1 <= max_chars:
            current = (current + "\n" + para) if current else para
        else:
            if current:
                chunks.append(current)
            if len(para) > max_chars:
                for i in range(0, len(para), max_chars):
                    chunks.append(para[i:i + max_chars])
                current = ""
            else:
                current = para

    if current:
        chunks.append(current)

    return chunks if chunks else [""]


def translate_text(
    text: str,
    direction: str,
    api_key: str,
    model: str,
    progress_bar,
    status_text,
) -> str:
    """Call DeepSeek API to translate text, preserving ⟨Tn⟩ placeholders."""
    client = OpenAI(
        api_key=api_key,
        base_url="https://api.deepseek.com",
    )

    if direction == "cn2en":
        lang_instruction = "Translate the following Chinese text to English."
    elif direction == "en2cn":
        lang_instruction = "Translate the following English text to Chinese."
    else:
        lang_instruction = (
            "Translate the following text to the other language "
            "(Chinese → English or English → Chinese, whichever applies)."
        )

    system_prompt = (
        f"{lang_instruction}\n\n"
        "CRITICAL RULES:\n"
        "1. Preserve ALL placeholders like ⟨T0⟩, ⟨T1⟩, ⟨T123⟩ — "
        "these are special tokens. Keep them EXACTLY as-is at the position "
        "where they appear. Do NOT translate, modify, split, or remove them.\n"
        "2. The surrounding text should read as natural, fluent prose in the "
        "target language. Adjust grammar, word order, prepositions, and "
        "articles as needed to accommodate where placeholders sit.\n"
        "3. Preserve paragraph breaks from the original.\n"
        "4. Output ONLY the translated text. No markdown, no HTML, no "
        "commentary, no notes, no explanations."
    )

    chunks = chunk_text(text)
    total_chunks = len(chunks)

    translated_chunks = []
    for i, chunk in enumerate(chunks):
        progress_bar.progress(
            (i / total_chunks) * 0.9,
            text=f"翻译中… ({i + 1}/{total_chunks})",
        )
        status_text.text(f"正在翻译第 {i + 1}/{total_chunks} 段…")

        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": chunk},
            ],
            max_tokens=8192,
            temperature=0.3,
        )
        translated_chunks.append(response.choices[0].message.content)

    return "\n".join(translated_chunks)


# ---------------------------------------------------------------------------
# Translation execution
# ---------------------------------------------------------------------------
if translate_clicked:
    # --- Step 1: Extract text ---
    with st.status("📄 提取文档文本…", expanded=True) as status:
        st.write("正在读取文档…")
        source_text = extract_text(doc_file.getvalue(), doc_file.name)
        st.write(f"✅ 提取完成，共 **{len(source_text):,}** 字符")

        if not source_text.strip():
            st.error("文档内容为空，请检查文件。")
            st.stop()

        # --- Step 2: Load glossary & protect terms ---
        use_glossary = False
        glossary_data = None
        if glossary_file:
            st.write("📖 加载用户术语表…")
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                tmp.write(glossary_file.getvalue())
                glossary_path = tmp.name
            glossary_data = load_glossary(glossary_path)
            os.unlink(glossary_path)
            use_glossary = True
        else:
            st.write("📖 加载内置术语表…")
            glossary_data = _decode_embedded_glossary()
            use_glossary = True

        protected_data = None
        text_to_translate = source_text

        if use_glossary:
            st.write(f"🔍 术语表共 **{glossary_data['count']}** 条术语，正在匹配…")
            protected_data = protect_terms(source_text, glossary_data, direction)
            text_to_translate = protected_data["protected_text"]
            st.write(
                f"✅ 匹配到 **{protected_data['matched_count']}** 条术语，已替换为占位符"
            )

        # --- Step 3: Translate ---
        status.update(label="🤖 AI 翻译中…", state="running")
        progress_bar = st.progress(0, text="准备翻译…")
        status_text = st.empty()

        translated = translate_text(
            text_to_translate,
            direction,
            effective_key.strip(),
            model,
            progress_bar,
            status_text,
        )
        progress_bar.progress(0.95, text="翻译完成，正在恢复术语…")
        status_text.text("翻译完成！")

        # --- Step 4: Restore terms ---
        if protected_data:
            st.write("🔄 恢复术语…")
            final_text = restore_terms(translated, protected_data)
            st.write(f"✅ 已恢复 **{protected_data['matched_count']}** 条术语")
        else:
            final_text = translated

        progress_bar.progress(1.0, text="✅ 完成！")
        progress_bar.empty()
        status_text.empty()

        status.update(label="✅ 翻译完成！", state="complete")

        st.toast("🎉 翻译完成！请查看结果并下载译文。", icon="✅")

        # --- Step 5: Store in session state for persistent display ---
        st.session_state.translation_result = {
            "source_text": source_text,
            "final_text": final_text,
            "direction": direction,
            "protected_data": protected_data,
            "glossary_data": glossary_data,
            "doc_filename": doc_file.name,
        }

# ---------------------------------------------------------------------------
# Display results (persisted in session state)
# ---------------------------------------------------------------------------
if "translation_result" in st.session_state:
    result = st.session_state.translation_result
    source_text = result["source_text"]
    final_text = result["final_text"]
    direction = result["direction"]
    protected_data = result["protected_data"]
    glossary_data = result["glossary_data"]
    doc_filename = result["doc_filename"]

    st.divider()
    st.markdown("## 📋 翻译结果")

    # --- Stats row in cards ---
    stat_col1, stat_col2, stat_col3, stat_col4 = st.columns(4)
    with stat_col1:
        st.metric("📊 源文字符", f"{len(source_text):,}")
    with stat_col2:
        st.metric("📝 译文字符", f"{len(final_text):,}")
    with stat_col3:
        st.metric(
            "🌍 翻译方向",
            {"cn2en": "中→英", "en2cn": "英→中"}.get(direction, direction),
        )
    with stat_col4:
        matched = protected_data["matched_count"] if protected_data else 0
        st.metric("🏷️ 匹配术语", f"{matched} 条")

    # --- Tabs ---
    tab1, tab2, tab3 = st.tabs(["📝 译文预览", "📄 原文对照", "📊 术语详情"])

    with tab1:
        st.text_area(
            "翻译结果",
            value=final_text,
            height=420,
            label_visibility="collapsed",
            key="result_preview",
        )

    with tab2:
        col_a, col_b = st.columns(2)
        with col_a:
            st.caption("**📌 原文**")
            st.text_area(
                "原文",
                value=source_text[:5000]
                + ("…" if len(source_text) > 5000 else ""),
                height=420,
                label_visibility="collapsed",
                key="source_preview",
            )
        with col_b:
            st.caption("**✅ 译文**")
            st.text_area(
                "译文",
                value=final_text[:5000]
                + ("…" if len(final_text) > 5000 else ""),
                height=420,
                label_visibility="collapsed",
                key="translated_preview",
            )

    with tab3:
        if protected_data and protected_data["matched_terms"]:
            st.dataframe(
                [
                    {"源术语": t["source"], "目标翻译": t["target"]}
                    for t in protected_data["matched_terms"]
                ],
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("未使用术语表，无法显示术语详情。")

    # --- Download button ---
    st.divider()
    col_dl1, col_dl2, col_dl3 = st.columns([1, 3, 1])
    with col_dl2:
        output_filename = Path(doc_filename).stem + "_translated.txt"
        st.download_button(
            label="📥 下载译文 (.txt)",
            data=final_text.encode("utf-8"),
            file_name=output_filename,
            mime="text/plain",
            type="primary",
            use_container_width=True,
        )

    # --- Matched terms log (collapsed) ---
    if protected_data:
        with st.expander("🔍 术语匹配详情"):
            st.json(
                {
                    "direction": direction,
                    "total_glossary_terms": glossary_data["count"],
                    "matched_count": protected_data["matched_count"],
                    "matched_terms": [
                        {"source": t["source"], "target": t["target"]}
                        for t in protected_data["matched_terms"]
                    ],
                }
            )

# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown("""
<div class="app-footer">
    术语表辅助翻译工具 · Powered by DeepSeek
</div>
""", unsafe_allow_html=True)
