import React from 'react'

export default function Header() {
    const lastExportDate="01/06/2024";
  return (
    <div className='mb-6'>
        <h3 className='text-gray-400'>Production Functionality Matrix</h3>
        <h3 className='text-gray-400'>v2</h3>
        <h3 className='text-gray-400'>Last Export: {lastExportDate}</h3>
    </div>
  )
}
