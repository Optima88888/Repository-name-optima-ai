<script id="mkt-price-detail-final-js">
(function(){
  'use strict';

  var plans={
    monthly:{title:'Gói 1 tháng',price:'159.000đ',amount:'159000',desc:'Phù hợp người mới bắt đầu, shop nhỏ cần đăng bài, tạo content và quản lý Fanpage cơ bản.',benefits:['Đăng bài Facebook','Quản lý Fanpage','Quản lý Group','AI Comment','Tạo content cơ bản','Token Manager','Hỗ trợ kích hoạt theo ID thiết bị']},
    quarterly:{title:'Gói 3 tháng',price:'359.000đ',amount:'359000',desc:'Tối ưu cho shop đang bán hàng cần dùng ổn định hơn và tiết kiệm hơn gói tháng.',benefits:['Toàn bộ gói 1 tháng','AI Messenger','CRM Kanban cơ bản','Kịch bản inbox','Lịch đăng nâng cao','Báo cáo cơ bản','Ưu tiên hỗ trợ']},
    halfyear:{title:'Gói 6 tháng',price:'559.000đ',amount:'559000',desc:'Phù hợp shop cần CRM, chăm sóc khách và tối ưu quy trình bán hàng.',benefits:['Toàn bộ gói 3 tháng','CRM Pro','AI Sales Bot','Comment Manager','Auto Tag khách hàng','Quản lý khách hàng','Chuyển khách sang CRM']},
    yearly:{title:'Gói 1 năm',price:'859.000đ',amount:'859000',desc:'Gói phổ biến nhất cho nhà bán hàng muốn dùng đầy đủ công cụ AI Marketing trong 1 năm.',benefits:['Toàn bộ gói 6 tháng','AI Marketing Director','AI Ads Chuyên Gia','Kho Content Premium','Automation Marketing','Export báo cáo','Ưu tiên xử lý']},
    sellerpro:{title:'Gói nhà bán hàng chuyên nghiệp',price:'1.959.000đ',amount:'1959000',desc:'Gói cao nhất cho nhà bán hàng chuyên nghiệp, mở toàn bộ hệ thống sau khi admin duyệt.',benefits:['Toàn bộ tính năng Premium','AI Image Center','AI Video Center','AI Voice Studio','Dashboard Enterprise','Export PDF / Excel','Backup Database','Ưu tiên hỗ trợ VIP']}
  };

  plans.lifetime = plans.sellerpro;

  function norm(t){
    return String(t||'')
      .toLowerCase()
      .replace(/\s+/g,' ')
      .replace(/vnđ|vnd|đ/g,'')
      .trim();
  }

  function keyFromText(t){
    t = norm(t);

    if(t.includes('1959000') || t.includes('1.959') || t.includes('1959') || t.includes('nhà bán') || t.includes('seller') || t.includes('chuyên nghiệp') || t.includes('trọn đời')) return 'sellerpro';
    if(t.includes('859000') || t.includes('859') || t.includes('1 năm') || t.includes('12 tháng')) return 'yearly';
    if(t.includes('559000') || t.includes('559') || t.includes('6 tháng')) return 'halfyear';
    if(t.includes('359000') || t.includes('359') || t.includes('3 tháng')) return 'quarterly';
    if(t.includes('159000') || t.includes('159') || t.includes('1 tháng')) return 'monthly';

    return '';
  }

  function findPlanKey(btn){
    var key =
      btn.getAttribute('data-plan') ||
      btn.getAttribute('data-package') ||
      btn.getAttribute('data-key') ||
      btn.dataset.plan ||
      btn.dataset.package ||
      btn.dataset.key ||
      '';

    key = norm(key);

    if(plans[key]) return key;

    var card = btn.closest('.pricing-card,.price-card,.plan-card,.premium-card,.package-card,.mkt-plan-card,.v2-price-card,div');
    var text = (btn.innerText || '') + ' ' + (card ? card.innerText : '') + ' ' + document.body.innerText;

    return keyFromText(text) || 'monthly';
  }

  document.addEventListener('click', function(e){
    var btn = e.target.closest('button,a');
    if(!btn) return;

    var text = norm(btn.innerText || btn.textContent || '');

    if(
      text.includes('nâng cấp') ||
      text.includes('thanh toán') ||
      text.includes('chọn gói') ||
      text.includes('mua gói')
    ){
      var key = findPlanKey(btn);
      window.MKT_SELECTED_PLAN_KEY = key;
      localStorage.setItem('mkt_selected_plan_key', key);
    }
  }, true);

})();
</script>