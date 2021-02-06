Hướng dẫn cách làm sổ sách kế toán trên Excel
=============================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-cach-lam-so-sach-ke-toan-tren-excel/)Hướng dẫn cách làm sổ sách kế toán trên Excel; Cách sử dụng hàm để hạch toán nghiệp vụ, để làm sổ sách, lập sổ chi tiết các tài khoản; Cách lập Báo cáo tài chính trên file Excel… Chú ý:  – Bài viết này HDDTVN xin hướng dẫn làm sổ sách kế toán trên …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



Hướng dẫn cách làm sổ sách kế toán trên Excel; Cách sử dụng hàm để hạch toán nghiệp vụ, để làm sổ sách, lập sổ chi tiết các tài khoản; Cách lập Báo cáo tài chính trên file Excel…
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



**Chú ý:**  

 – Bài viết này HDDTVN xin hướng dẫn làm sổ sách kế toán trên file Excel mẫu do HDDTVN thiết kế.  

 – Mẫu các bạn có thể tải về tại đây nhé: **[Mẫu sổ sách kế toán trên Excel](# "mẫu sổ sách kế toán trên excel").**



  

—————————————————————–  

  
**1. Tìm hiểu các hàm thường sử dụng trong Excel kế toán:**  

   

Các bạn muốn làm tốt công việc của người kế toán trên Excel thì việc đầu tiên các bạn cần quan tâm đó là: Hiểu rõ tác dụng của các hàm:  

   

**a. Tác dụng của HÀM SUMIF:**  

– Kết chuyển các bút toán cuối kỳ  

– Tổng hợp số liệu từ NKC lên Phát sinh Nợ, Phát sinh Có trên Bảng cân đối số phát sinh tháng và năm  

-Tổng hợp số liệu từ PNK, PXK lên “**Bảng NHập Xuất Tồn**“  

– Tổng hợp số liệu từ NKC lên cột PS Nợ, PS Có của “**Bảng tổng hợp phải thu, phải trả khách hàng**”


**Chú ý:** Với các bút toán kết chuyển cuối kỳ thì điều kiện cần tính có thể bấm trực tiếp vào ô chứa nó hoặc gõ trực tiếp tài khoản cần kết chuyển vào công thức



Ví dụ: = **SUMIF( $E15:$E180, 5111,$H$15:$H$180)**
   

**b. Tác dụng của Hàm VLOOKUP:**  

– Tìm đơn giá Xuất kho từ bên Bảng Nhập Xuất Tồn về Phiếu Xuất kho.  

– Tìm Mã TK, Tên TK từ Danh mục tài khoản về bảng CĐPS, về Sổ 131, 331…  

– Tìm Mã hàng hoá, tên hàng hoá từ Danh mục hàng hoá về Bảng Nhập Xuất Tồn  

– TÌm số dư của đầu tháng N căn cứ vao cột số dư của tháng N-1  

– Tìm số Khấu hao (Phân bổ) luỹ kế từ kỳ trước, căn cứ vào Giá trị khấu hao (phân bổ) luỹ kế (của bảng 242, 214)


**Ví dụ:** Đi tìm Đơn giá xuất kho của mã hàng GEN\_ASG24R từ bảng Nhập xuất tồn về Phiếu Xuất kho.



Ví dụ **=VLOOKUP(F11,$C$20:$D$22,2,0)**

Ngoài ra còn 1 số hàm nữa, chi tiết tại đây nhé: **[Các hàm thường dùng trong kế toán Excel](# "các hàm thường dùng trong kế toán excel")**



  

 ————————————————————  

  
**2. Các công việc đầu năm cần phải làm khi làm sổ sách:**  

   

– Những DN đã và đang hoạt động thì đầu năm các bạn phải chuyển số dư cuối năm trước sang đầu năm nay, cụ thể như sau:  

– Vào số dư đầu kỳ “Bảng cân đối phát sinh tháng”  

– Vào số dư dầu kỳ các Sổ chi tiết tài khoản 242, 211, 131,  

– Vào Bảng tổng hợp Nhập Xuất Tồn, và các Sổ khác (nếu có)  

– Chuyển lãi (lỗ) năm nay về năm trước (Căn cứ vào số dư đầu kỳ TK 4212 trên Bảng CĐTK để chuyển). Việc thực hiện này được hạch toán trên Nhật ký chung và chỉ thực hiện 1 lần trong năm, vào thời điểm đầu năm.





  

————————————————————–
   

**3. Các công việc trong tháng cần phải làm:**


– Hiện nay đa phần DN lựa chọn ghi sổ theo hình thức Nhật ký Chung, nên Công ty HDDTVN xin hướng dẫn cách ghi sổ nhật ký chung trên Excel.  

   

**Chú ý:** Trong quá trình hạch toán ghi sổ bạn phải làm theo nguyên tắc đó là: **Đồng nhất về tài khoản và đồng nhất về mã hàng hóa.**  

 




   

**Cách ghi sổ khi có phát sinh tại Doanh nghiệp:**


**a. Khi mua hàng**:


**+) Khi có phát sinh** thêm Khách hàng hoặc nhà cung cấp mới:  

    – Thì các bạn khai báo chi tiết đối tượng KH hoặc NCC mới bên bảng Danh mục tài khoản và đặt mã Tài khoản (Mã khách hàng) cho KH/NCC đó, đồng thời định khoản chi tiết bên NKC theo mã TK mới khai báo.


Ví dụ: Phải thu của Công ty HDDTVN (là khách hàng mới):  

**Bước 1:** Sang DMTK khai chi tiết khách hàng – Công ty HDDTVN với mã Khách hàng mới là: 1311 hoặc 131TU (Khai báo phía dưới Tk 131)  

    -> *(Việc khai báo mã TK như thế nào là tuỳ vào yêu cầu quản trị của bạn, sao cho dễ kiêm tra đối chiếu)*


**Bước 2:** Hạch toán bên NKC theo mã TK (Mã KH) đã khai báo cho Công ty HDDTVN là 1311 hoặc 131TU


    – Nếu không phát sinh Khách hàng mới thì khi gặp các nghiệp vụ liên quan đến TK 131  và TK 331, ta quay lại Danh mục TK để lấy Mã Khách hàng đã có và định khoản trên NKC.


**+) Trường hợp** phát sinh mới Công cụ dụng cụ hoặc TSCĐ (tức liên quan đến TK 242, 214 )  

    – Sau khi định khoản trên NKC phải sang bảng phân bổ 242, 214 để khai báo thêm công cụ dụng cụ hoặc tài sản này vào bảng và tính ra số cần phân bổ trong kỳ hoặc số cần trích khấu hao trong kỳ.


**+) Khi mua hàng hóa:**  

**Bước 1:** Bên Nhật ký chung không phải khai chi tiết từng mặt hàng mua vào, chỉ hạch toán chung vào TK 156 tổng số tiền ở dòng “Cồng tiền hàng” trên hoá đơn mua vào  

   

**– Bước 2:** Đồng thời về Phiếu nhập kho, khia báo chi tiết từng mặt hàng mua theo hoá đơn vào phiếu nhập kho:  

– Nếu mặt hàng mua vào đã có tên trong Danh mục hàng háo thì quay về DM hàng hoá để lấy Mã hàng, tên hàng cho hàng hoá đó và thực hiện kê nhập.  

– Nếu mặt hàng mùa vào là hàng mới thì phải đặt Mã hàng cho từng mặt hàng trên DMHH sau đó thực hiện kê nhập trên PNK theo mã hàng đã khai báo


**Bước 3:** Nếu phát sinh chi phí (vận chuyển, bốc dỡ, lưu kho…) cho việc mua hàng thì Đơn giá nhập kho là đơn giá đã bao gồm chi phí. Khi đó phải phân bổ chi phí mau hàng cho từng mặt hàng như sau:  

*(Có thể lập bảng tính riêng cho việc phân bổ chi phí)*






**Chi phí của mặt hàng A**

 = 

**Tổng chi phí**

**X**

**số lượng ( hoặc thành tiền) mặt hàng A**



**Tổng số lượng ( hoặc tổng thành tiền Của lô hàng )**








**Chi phí đơn vị Mặt hàng A**

**=**

**Tổng chi phí của mặt hàng A**



**Tổng số lượng của mặt hàng A**








**Đơn giá nhập kho Mặt hàng A**

**=**

**Đơn giá mua của Mặt hàng A**

**+**

**Chi phí đơn vị Của mặt hàng A**




**b. Khi bán hàng hóa:**  

   

**Bước 1:** Bên Nhật ký chung không phải khai chi tiết từng mặt hàng bán ra, chỉ hạch toán chung vào TK 5111 tổng số tiền ở dòng “CỘng tiền hàng” trên hoá đơn bán ra.


**Bước 2:** Đồng thời về Phiếu Xuất kho, khai báo chi tiết từng mặt hàng bán ra theo Hoá đơn vào Phiếu XK.  

– Để lấy được Mã hàng xuất kho, ta quay về Danh mục hàng hoá để lấy.  

– Không hạch toán bút toán Giá vốn hàng bán: Vì CÔng ty áp dụng phương pháp tính gái xuất kho là phương pháp “**Bình quân cuối kỳ**“, nên cuối tháng mới thực hiện bút toán này để tập hợp giá vốn hàng bán trong kỳ.


**Chú ý:**  

– Khi vào bảng kê xuất kho thì chỉ vào số lượng, chưa có đơn gá xuất kho vì đơn giá cuối kỳ mới tính được bên Bảng Nhập Xuất TỒn kho  

– Khi tính được Đơn giá bên bảng Nhập – Xuất – Tồn thì sử dụng hàm VLOOKUP tìm đơn giá xuất kho từ bảng Nhập – XUất – TỒn về PXK.


**c. Khi cần nhập kho hoặc xuất kho:**  

– Các bạn vào phiếu nhập kho, xuất kho (cũng làm theo như phần trên nhé)



  

 —————————————————————–  

  
**4. Công việc cuối tháng cần làm:**  

   

– Thực hiện các bút toán kết chuyển cuối tháng như: Kết chuyển tiền lương, trích khấu hao TSCĐ, phân bổ chi phí, kết chuyển thuế, kết chuyển doanh thu, chi phí…  

 





  

 ————————————————————————-  

  
**5. Lập các bảng biểu tháng:**  

   

– Lập bảng tổng hợp Nhập – Xuất – Tồn kho.  

– Lập bảng Phân bổ Chi phí trả trước, khấu hao TSCĐ.  

– Lập bảng cân đối phát sinh tháng.



[**Cách lập các bảng biểu cuối tháng trên Excel**](# "cách lập các bảng biểu cuối tháng trên excel")

  

 ————————————————————————-  

  
**6. Kiểm tra số liệu trên Bảng cân đối phát sinh.**  

   

– Trên CĐPS thì tổng phát sinh bên Nợ phải bằng tổng phát sinh bên Có  

– Tổng PS Nợ trên CĐPS bằng tổng PS Nợ trên NKC  

– Tổng PS Có trên CĐPS bằng Tổng PS Có trên NKC  

– Các tài khoản loại 1 và loại 2 không có số dư bên Có. Trừ một số tài khoản như 131, 214,..  

– Các tài khoản loại 3 và loại 4 không có số dư bên Nợ, trừ một số tài khoản như 331, 3331, 421,..  

– Các tài khoản loại 5 đến loại 9 cuối kỳ không có số dư.  

– TK 112 phải khớp với Sổ phụ ngân hàng,  

– TK 133, 3331 phải khớp với chỉ tiêu trên tờ khai thuế  

– TK 156 phải khớp với dòng tổng cộng trên Báo cáo NXT kho  

– TK 242 phải khớp với dòng tổng cộng trên bảng phân bổ 242  

– TK 211 , 214 phải khớp với dòng tổng cộng trên Bảng khấu hao 211



  

 ———————————————————————-  

  
**7. Lập sổ chi tiết các Tài khoản cuối kỳ:**  

   

**a. Lập bảng Tổng hợp phải thu khách hàng – TK 131:**


– Cột Mã KH và Tên KH, sử dụng hàm IF kết hợp với hàm LEFT tìm ở DMTK về


– Cột mã khách hàng: = IF(LEFT(DMTK!A14,3)=”131”,DMTK!A14,””) (địa chỉ ô A 14 là TK chi tiết đầu tiên của TK 131). Hoặc dùng VLOOKUP tìm DMTK về


– Cột tên khách hàng; = IF(A11=””,””,VLOOKUP(A11,DMTK!$A$14:$E$150,2,0))


– Dư Nợ và Dư có đầu kỳ: Dùng hàm VLOOKUP tìm ở CĐPS tháng về.


– Cột Dư NỢ đầu kỳ: = VLOOKUP($A11,’CDSPS-thang’!$A$9:$H$150,3,0)


– Cột Dư Có đầu kỳ: = VLOOKUP($A11,’CDSPS-thang’!$A$9:$H$150,4,0)


– Cột số phát sinh Nợ và phát sinh Có, sử dụng hàm SUMIF tập hợp từ NKC về.


– Cột số phát sinh Nợ:  

= SUMIF(NKC!$E$13:$E$607.’So cong no TK 131’!$A11,NKC!$G$13:$G$607)


– Cột số phát sinh Có:  

= SUMIF(NKC!$E$13:$E$607,’So cong no TK 131’$A11,NKC!$H$13:$H$607)


– Cột dư Nợ và dư Có cuối kỳ, dùng hàm Max  

– Cột dư Nợ cuối kỳ:  = MAX ( C11+E11-D11-F11,0)  

– Cột dư Có cuối kỳ:   = MAX( D11+F11–C11-E11,0)  

   

**b. Lập Bảng tổng hợp Phải trả khách hàng – TK 331**  

– Cách làm tương tự như bảng tổng hợp TK 131  

   

**c. Lập sổ quỹ tiền mặt và sổ tiền gửi ngân hàng;**


– Riêng sổ quỹ tiền mặt và sổ Tiền gửi ngân hàng chúng ta không thể chuyển sổ trên NKC mà phải tính riêng 2 sổ này, vì 2 loại sổ này có mẫu sổ khác so với các sổ chi tiết TK, sổ tổng hợp TK khác  

   

**+) Cách lập sổ Quỹ tiền mặt: ( Dữ liệu lấy từ sổ Nhật Ký Chung)**  

– Cách lập công thức cho từng cột như sau: Trên sổ quỹ tiền TM, xây dựng thêm 3 ô: Tháng báo cáo; Tài khoản báo cáo ( là TK 1111); Nối tháng và TK cáo cáo.  

– Ô nối tháng và TK báo cáo = K6&”;”&L6 ( dùng tính số dư dầu kỳ theo từng tháng)  

– Cột ngày tháng: = IF($L$6=NKC!$E13,NKC!A13,””) ( L6: ô TK báo cáo)  

– Cột Diễn giải: = IF($L$6=NKC!$E13,NKC!D13,””)  

– Cột Tài khoản đối ứng: = IF($L$6=NKC!$E13,NKC!F13,””)  

– Cột thu:   = IF($L$6=NKC!$E13,NKC!G13,””)  

– Cột Chi:  = IF($L$6=NKC!$E13,NKC!H13,””)  

– Cột số phiếu thu: = IF(H11<=0,””,COUNTIF(H$11:H11,”>”)) ( H: cột tiền thu)  

– Cột số phiếu chi:  = IF(I11<=0,””,COUNTIF(I$11:I11,”>”))   ( I: Cột tiền chi)  

– Dòng số dư dầu kkỳ dùng hàm SUMIF lấy trên bảng CDPS chi tiết của từng tháng.  

   

*Để tính được số dư đầu kỳ của từng tháng trên Sổ quỹ TM thì ta phải xây dựng bên phải bảng “Cân đối phát sinh tháng” của các tháng thêm 2 cột:*  

– Cột BC: Gõ số tháng tại dòng tương ứng với TK 111 của Bảng CĐPS và coppy cho những dòng tiếp theo của tháng đó ( làm cho tất cả các tháng).  

– Cột “ Nối tháng và TK báo cáo” : =I9&”;”&A9 ( Là dãy điều kiện cho hàm SUMIF)  

   

*Sau đó dùng hàm SUMIF để tính ra số dư đầu kỳ trên sổ Quỹ TM;*   

= SUMIF(‘CDPS-thang’!$J$9:$J$450,’So quy’!$M$6,’CDPS-thang’!$C$9:$C$450)  

– Cột tồn tiền cuối ngày dùng hàm Subtotal:  

Cú pháp hàm: = $J$9+Subtotal(9,H$11:H11)-Subtotal(9,I$11:I11)  

– Dòng cộng số phát sinh : Dùng hàm subtotal  

– Dòng số dư cuối kỳ: Dùng công thức đơn giản như sau:  

Dư cuối kỳ = Tồn đầu kỳ + TỔng thu – TỔng chi.  

( Sổ quỹ TM được lập cho cả kỳ kế toán, Bạn muốn xem tháng nào thì lọc tháng đó lên, Cụ thể có ở phần in sổ)  

   

**+) Lập sổ tiền gửi ngân hàng:**  

– Cách làm tương tự như sổ quỹ tiền mặt. Nhưng cột số hiệu và Ngày tháng chứng từ thì công thức tương tự như cột Ngày tháng ghi sổ.



  

——————————————————————-  

  

**8. Cách lập báo cáo tài chính cuối năm:**


**a. Bảng cân đối kế toán (Báo cáo tình hình tài chính)**



*(Để bảng cân đối kế toán đúng thì Tổng Tài Sản phải bằng tổng Nguồn Vốn)*


**Cách làm:**  

 – Cột số năm trước: Căn cứ vào Cột năm nay của “Bảng Cân Đối Kế toán” hoặc “Báo cáo tình hình tài chính” Năm trước.  

 – Cột số năm nay: Chuyển số liệu của các TK từ loại 1 đến loại 4 ( phần số dư cuối kỳ ) trên bảng CĐPS năm và ghép vào từng chỉ tiêu tương ứng trên Bảng CĐKT. Ví dụ : Chỉ tiêu [110] – “ Tiền và các khoản tương đương tiền “ bằng (=) Số dư Nợ cuối kỳ của các tài khoản 111 + TK 112 + TK 121 ( đối với các khoản đầu tư ngắn hạn có thời hạn dưới 3 tháng ).  

 – Riêng đối với các chỉ têu liên quan đến khách hàng và nhà cung cấp ( người bán ) thì căn cứ vào Bảng Tổng Hợp TK 131, 331.



 (theo TT 133)



**[Cách lập bảng cân đối kế toán](# "cách lập bảng cân đối kế toán")** (theo TT 200)



  

**b. Bảng báo cáo kết quả kinh doanh** (Bảng này lập cho thời kỳ – là tập hợp kết quả kinh doanh của một kỳ, thời kỳ ở đây có thể là tháng, quý Năm tuỳ theo mục đích quản trị của Doanh Nghiệp và là năm đối với cơ quan  thuế )


**Cách làm:**  

 – Cột số năm trước: Căn cứ vào cột ngăm ngay của “Báo cáo kết quả kinh doanh “ năm ttrước  

 – Cột số năm nay : CHuyển số liệu từ Bảng CĐPS năm của các TK từ loại 5 đến loại 8 ( phần số phát sinh ) và ghép vào từng chỉ tiêu tương ứng trên Báo cáo KQKD.  

**Ví dụ:** Chỉ tiêu [01] – “ Doanh thu bán hàng và cung cấp dịch vụ “ bằng (= ) Tổng số phát sinh Tk 511 trên bảng CĐPS năm.



 (theo TT 133)  

**[Hướng dẫn lập báo kết quả kinh doanh](# "hướng dẫn lập báo cáo kết quả kinh doanh")** (theo TT 200)



  

**c. Báo cáo lưu chuyển tiền tệ** (Thể hiện dòng ra và dòng vào của tiền trong Doanh Nghiệp, để BCLCTT đúng thì chỉ tiêu (70) trên LCTT phải bằng chỉ tiêu (110) trên bảng CĐKT)  

 Cách làm:  

 – Cột số năm trước: Căn cứ vào Cột năm nay của “Báo cáo lưu chuyển tiền tệ” năm trước.  

 – Cột Số năm nay : Căn cứ vào sổ Quỹ tiền mặt và Tiền gửi Ngân hàng, hoặc căn cứ vào số phát sinh TK tiền mặt, TK tiền gửi ngân hàgn trên NKC.  

    

**Nếu căn cứ vào Sổ Quỹ Tiền Mặt và Tiền gửi Ngân hàng :**  

 – Trên sổ quỹ TM, tính tổng số phát sinh của cả kỳ kế toán tại cột thu, chi bằng hàm subtotal  

 – Đặt lọc cho Sổ quỹ TM ( lưu ý không lọc cácc chỉ tiêu đè có định )  

 – Trên cột TK đối ứng lọc lần lượt từng TK đối ứng vừa lọc và bên cột Diễn giải sẽ xuất hiện nội dung của nghiệp vụ. Nội dung này tương ứng với từng chỉ tiêu nào trên “ BC lưu chuyển tiền tệ “ thì mang số tiền tổng cộng về đúng chỉ tiêu đó trên “ BC lưu chuyển tiền tệ “. Nếu có nhiều nội dung chung cho một chỉ tiêu thì thực hiện cộng nối tiếp vào sổ đã có. Nếu nội dung lọc lên màk hông biết đưa vào chỉ tiêu nào thì đưa vào thu khác hoặc chi khác.  

*(Thực hiện tương tự như sổ ngân hàng )*  

    

**Nếu căn cứ từ Nhật Ký Chung:**  

 – TÍnh tổng cộng phát sinh của cả kỳ kế toán trên NKC bằng hàm subtotal  

 – Đặt lọc cho sổ NKC ( Lưu ý : không lọc các tiêu đề cố định )  

 – Trên NKC, ở Cột TK nợ/ TK Có các bạn lọc lên TK Tiền Mặt, sau đó lọc tiếp lần lượt từng TK đối ứng bên Cột TK đối ứng. Khi Đó hàm subtotal sẽ tính tổng số tiền của TK đối ứng vừa lọc và bên cột Diễn giải sẽ xuất hiện nội dung của nghiệp vụ..  

*(Làm tương tự các phần tiếp theo như khi căn cứ vào sổ Quỹ tiền mặt)*



 (theo TT 133)  

**[Hướng dẫn lập báo cáo lưu chuyển tiền tệ](# "hướng dẫn lập báo cáo lưu chuyển tiền tệ")** (theo TT 200)



  

**d. Thuyết minh báo cáo tài chính:**  

 – Là báo cáo giải thích thêm cho các biểu: bảng CĐKT, báo cáo KQKD, Báo cáo lưu chuyển tiền tệ. Các bạn căn cứ vào Bảng Cân Đối Kế Toán, Báo cáo KQKD, Báo cáo LCTT, Bảng cân đối số phát sinh năm, Bảng trích Khấu hao TSCĐ ( trường hợp thuyết minh cho phần TSCĐ) và các sổ sách liên quan để lập.



 (theo TT 133)  

**[Hướng dẫn lập thuyết minh báo cáo tài chính](# "hướng dẫn lập thuyết minh báo cáo tài chính")** (theo TT 200)



  

**đ. Lập bảng Cân đối phát sinh năm:**  
  

– Có 2 dạng bảng cân đối phát sinh năm:  
  

  

**+) Dạng bảng chi tiết:** thì lập tương tự như Cân dối phát sinh tháng, với danh mục tài khoản là danh mục chi tiết, số liệu tập hợp từ NKC của cả năm.  

**+) Dạng bảng tổng hợp:**  
  

– Bảng này là bảng tổng hợp, nên được lập cho tài khoản cấp 1 ( trừ 333)  
  

– Số liệu được tập hợp từ NKC của cả năm  
  

   
  

**Cách làm:**  
  

– Trên Nhật Ký chung. Xây dựng thêm cột TK cấp 1.  
  

– Sử dụng hàm LEFT cho cột TK cấp 1 để láy về TK cấp 1 từ Cột TK Nợ/ TK Có trên NKC.  
  

– Cột mã TK, tên TK: Dùng hàm VLOOKUP hoặc Coppy từ DMTK về, sau đó xoá hết TK chi itết( trừ các TK chi tiết của TK 333 )  
  

– Cột dư Nợ và dư Có đầu kỳ: Dùng hàm VLOOKUP tìm ở CĐPS tháng 1 về ( phần dư đầu kỳ)  
  

– Cột phát sinh Nợ, Phát sinh có: Dùng SUMIF tổng hợp ở nhật ký chung về ( dãy ô điều kiện vẫn là cột TK Nợ/TK có )  
  

– Cột dư Nợ, dư Có cuối kỳ: Dùng hàm MAX  
  

– Dòng tổng cộng dùng hàm SUBTOTAL (Lưu ý: Sử dụng hàm SUBTOTAL cho TK 333)



 (theo TT1 33)



—————————————————————————–




moreHướng dẫn làm excel kế toán; Cách lập sổ sách kế toán trên Excel sử dụng các hàm, lập sổ Nhật ký chung, sổ chi tiết các tài khoản, lập BCTC trên Excel…

