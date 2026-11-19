-- Web khảo sát nhu cầu CGKT & hỗ trợ chuyên môn 2027 – BV Phổi Hải Phòng
-- Chạy 1 lần trong Supabase > SQL Editor.

create table if not exists cau_hinh (
  id int primary key default 1,
  mo boolean not null default true,
  han_chot timestamptz,
  thong_bao text
);
create table if not exists khoa (
  id uuid primary key default gen_random_uuid(),
  ten text not null,
  thu_tu int default 0
);
create table if not exists don_vi (
  id uuid primary key default gen_random_uuid(),
  ten text not null,
  tuyen text not null check (tuyen in ('TW','TP')),
  ghi_chu text,
  duyet boolean not null default false
);
create table if not exists ky_thuat (
  id uuid primary key default gen_random_uuid(),
  ten text not null,
  chuyen_nganh text not null,
  duyet boolean not null default false,
  de_xuat_khoa uuid,
  created_at timestamptz default now()
);
create table if not exists lua_chon (
  id uuid primary key default gen_random_uuid(),
  ky_thuat_id uuid not null references ky_thuat(id) on delete cascade,
  don_vi_id uuid not null references don_vi(id) on delete cascade,
  doi_tuong text, thoi_gian text, goi text, hoc_phi text, nguon_tt text,
  duyet boolean not null default false
);
create table if not exists dang_ky (
  id uuid primary key default gen_random_uuid(),
  khoa_id uuid not null references khoa(id),
  loai text not null check (loai in ('nhan','cg','lp')),
  ky_thuat_id uuid references ky_thuat(id) on delete set null,
  lua_chon_id uuid references lua_chon(id) on delete set null,
  don_vi_id uuid references don_vi(id) on delete set null,
  ten_kt text, don_vi_nhan text, huong text, don_vi_khac text, chuyen_khoa text,
  hinh_thuc text, so_nguoi int, doi_tuong text, thoi_gian text, ghi_chu text,
  nguoi_khai text, sdt text,
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

-- Đang mở nhập liệu? (cờ "mo" + chưa quá hạn chót)
create or replace function dang_mo() returns boolean
language sql stable security definer set search_path = public as $$
  select coalesce((select mo and (han_chot is null or now() <= han_chot) from cau_hinh where id = 1), false)
$$;

alter table cau_hinh enable row level security;
alter table khoa     enable row level security;
alter table don_vi   enable row level security;
alter table ky_thuat enable row level security;
alter table lua_chon enable row level security;
alter table dang_ky  enable row level security;

-- Ai cũng xem được
do $$ declare t text; begin
  foreach t in array array['cau_hinh','khoa','don_vi','ky_thuat','lua_chon','dang_ky'] loop
    execute format('drop policy if exists xem on %I', t);
    execute format('create policy xem on %I for select to anon, authenticated using (true)', t);
    -- Quản trị (tài khoản đăng nhập) toàn quyền
    execute format('drop policy if exists admin_all on %I', t);
    execute format('create policy admin_all on %I for all to authenticated using (true) with check (true)', t);
  end loop;
end $$;

-- Khoa (không đăng nhập): chỉ ghi khi đang mở
drop policy if exists khoa_them on dang_ky;
drop policy if exists khoa_sua on dang_ky;
drop policy if exists khoa_xoa on dang_ky;
create policy khoa_them on dang_ky for insert to anon with check (dang_mo());
create policy khoa_sua  on dang_ky for update to anon using (dang_mo()) with check (dang_mo());
create policy khoa_xoa  on dang_ky for delete to anon using (dang_mo());

-- Khoa đề xuất kỹ thuật/đơn vị mới: chỉ thêm, ở trạng thái chờ duyệt
drop policy if exists de_xuat on ky_thuat;
drop policy if exists de_xuat on don_vi;
drop policy if exists de_xuat on lua_chon;
create policy de_xuat on ky_thuat for insert to anon with check (dang_mo() and duyet = false);
create policy de_xuat on don_vi   for insert to anon with check (dang_mo() and duyet = false);
create policy de_xuat on lua_chon for insert to anon with check (dang_mo() and duyet = false);

-- Sau khi chạy: Authentication > Users > Add user (email + mật khẩu quản trị),
-- và Authentication > Providers > Email: TẮT "Allow new users to sign up".
