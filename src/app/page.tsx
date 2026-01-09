'use client'
import React, { useEffect, useMemo, useRef, useState } from 'react';
import axios from 'axios';
import Image from 'next/image';
import done from '../assets/tick-green-icon.svg';
import loader from '../assets/tube-spinner.svg';
import * as XLSX from 'xlsx';
import { BsFiletypeXlsx } from 'react-icons/bs';
import { addToast } from '@heroui/toast';



export default function Home() {
  const [value, setValue] = useState<string>('');
  const [fetching, setFetching] = useState<boolean>(false);
  const [loadedAt, setLoadedAt] = useState<string>('-');
  const [columnAData, setColumnAData] = useState<number[]>([]);
  const [responseData, setResponseData] = useState<{
    name: string;
    taxNumber: string,
    date: string,
  }[]>([]);
  const ref = useRef<any>(null);
  const loadedAtRef = useRef<number>(0);

  const fetchData = (shouldSetLoading = true) => {
    if ((Date.now() - loadedAtRef.current) < 1000 * 60 * 30) {
      return;
    }
    setFetching(shouldSetLoading);
    const mainUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
    const tableId = process.env.NEXT_PUBLIC_GOOGLE_SHEETS_TABLE_ID;
    const apiKey = process.env.NEXT_PUBLIC_GOOGLE_SHEETS_API_KEY;

    axios.get(`${mainUrl}/${tableId}/values/echo-all!A2:G?key=${apiKey}`).then((res) => {
      const resFiltered = res.data.values.filter((el: any) => !!el.length).map((item: any) => ({
        name: `${item?.at(0)} ${item?.at(1)} ${item?.at(2)}`,
        taxNumber: item?.at(4),
        date: item?.at(5),
      })) || [];
      setFetching(false);
      setResponseData(resFiltered);
      if (resFiltered.length < 800) {
        showToast();
        setResponseData([]);
      }
      loadedAtRef.current = Date.now();
    }).catch(() => {
      setFetching(false);
      setResponseData([]);
      showToast();
    });

    axios.get(`${mainUrl}/${tableId}/values/updates!B1?key=${apiKey}`).then((res) => {
      setLoadedAt(res?.data?.values?.at(0)?.at(0))
    });
    const showToast = () => {
      addToast({
        title: 'Помилка імпорту Google Sheets',
        description: 'Будь ласка перезавантажте сторінку',
        color: 'danger',
        timeout: 1000 * 120,
      });
    }
  }

  useEffect(() => {
    fetchData();
    const handleFocus = () => {
      fetchData(false)
    }

    window.addEventListener('focus', handleFocus)
    return () => {
      window.removeEventListener('focus', handleFocus)
    }
  }, []);

  const list = useMemo(() => {
    if (value.trim() && value.length > 2) {
      return responseData.filter((el) => (el?.taxNumber?.startsWith('0') ? parseInt(el.taxNumber)?.toString() : el.taxNumber)?.startsWith(value?.startsWith('0') ? parseInt(value)?.toString() : value) || el?.name?.toLowerCase()?.startsWith(value.toLowerCase()))
    } else if (columnAData.length > 0) {
      return responseData.filter((el) => columnAData.includes(parseInt(el.taxNumber)));
    } else {
      return [];
    }
  }, [value, responseData, columnAData])

  const onClear = () => {
    setValue('');
  };


  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (evt: ProgressEvent<FileReader>) => {
      const binaryStr = evt.target?.result;
      if (typeof binaryStr !== 'string') return;

      const workbook = XLSX.read(binaryStr, {type: 'binary'});
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');

      const columnA: string[] = [];

      for (let row = range.s.r; row <= range.e.r; row++) {
        const cellAddress = {c: 0, r: row};
        const cellRef = XLSX.utils.encode_cell(cellAddress);
        const cell = worksheet[cellRef];
        if (cell && cell.v != null) {
          columnA.push(String(cell.v));
        }
      }

      setColumnAData(columnA?.map((e) => parseInt(e))?.filter((item) => !isNaN(item)));
    };

    reader.readAsBinaryString(file);
  };

  const exportToXlsx = () => {
    if (!list || list.length === 0) return;

    const rows = list.map((r) => ({
      Дата: r.date || '',
      Бенефіціар: r.name || '',
      ІПН: r.taxNumber || '',
    }));

    const ws = XLSX.utils.json_to_sheet(rows, {header: ['Дата', 'Бенефіціар', 'ІПН']});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Results');

    const ts = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
    const filename = `echo_check_${ts}.xlsx`;
    XLSX.writeFile(wb, filename);
  };


  return fetching ? (<div className="w-full h-[100vh] flex justify-center items-center">
    <Image src={loader} alt={'loading'} className="my-5"/>
  </div>) : (
    <div className="min-h-screen flex items-center  flex-col bg-gray-100">
      <div className="text-sm">
        Завантажено бенефіціарів: {responseData.length}
      </div>
      <div className="text-sm">
        Бази оновлено: {loadedAt}
      </div>
      <div className="text-center mb-4 px-10 text-xl max-w-[500px] pt-8">
        Перевірка реєстрації ECHO
      </div>
      <div className="relative max-w-[500px] w-[95%]">
        <input value={value} onChange={(e) => setValue(e.target.value.trim())} type="text"
               placeholder="Прізвище або ІПН"
               className="w-full text-2xl p-1 border-2 rounded-m border-gray-300 mb-1 pr-4"/>
        {!!value.trim() && <span onClick={onClear}
                                 className="h-fit absolute right-1 top-[3px] bottom-0 py-1.5 px-3 bg-red-300">X</span>}
      </div>
      <div className="hidden sm:block">
        <div className="text-sm text-center pt-4 mb-2">
          Для масової перевірки завантажте файл .XLSX <br/>який містить всі ІПН бенефіціарів в колонці А
        </div>
        <div className="flex gap-5 items-center">
          <input id="file-upload" ref={ref} type="file" accept=".xlsx, .xls" onChange={handleFileUpload}
                 className="border-b-2 text-center  px-3"/>
          <label htmlFor="file-upload"> <BsFiletypeXlsx size={30}/></label>
          {!!columnAData?.length && <button onClick={() => {
            setColumnAData([]);
            if (ref?.current) {
              ref.current.value = '';
            }
          }}>
            X
          </button>}
        </div>
      </div>

      {list?.length ? <div className="w-full max-w-[500px] py-5 flex flex-col">
        {
          list?.map((el, index: number) => (
            <div className="mb-2 border-b-2 border-gray-300 p-3" key={`${index}_${el.taxNumber}`}>
              <div>{el.name}</div>
              <div className="flex justify-between">
                <div>{el.taxNumber}</div>
                <div>{el.date}</div>
              </div>
            </div>
          ))
        }
        <button
          onClick={exportToXlsx}
          className="px-3 py-1 bg-blue-500 text-white rounded disabled:opacity-50 mx-auto my-5"
        >
          Скачати результат в форматі XLSX
        </button>
      </div> : ''}
      {((value && value.length > 2 && list?.length === 0) || (list?.length === 0 && columnAData?.length > 0)) &&
        <div className="flex flex-col justify-center items-center text-xl">
          <Image width={50} src={done} alt={'Ok'} className="my-5"/>
          Бенефіціара не знайдено
        </div>}
    </div>
  );
}
