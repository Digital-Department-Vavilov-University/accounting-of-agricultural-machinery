#pragma once

namespace Project1 {

	using namespace System::Windows::Forms;

	public ref struct RowData {
		DataGridViewCell^ beginDataCell;
		DataGridViewCell^ receivedDataCell;
		DataGridViewCell^ goneDataCell;
		DataGridViewCell^ resultCell;
	};
}