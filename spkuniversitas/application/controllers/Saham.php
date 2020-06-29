<?php
/**
 * Created by PhpStorm.
 * User: sankester
 * Date: 11/05/2017
 * Time: 15:42
 */

if (!defined('BASEPATH')) exit('No direct script access allowed');

class Saham extends MY_Controller
{
    public function __construct()
    {
        parent::__construct();
        $this->page->setTitle('Saham');
        $this->load->model('MKriteria');
        $this->load->model('MSubKriteria');
        $this->load->model('MSaham');
        $this->load->model('MNilai');
        $this->page->setLoadJs('assets/js/saham');
    }

    public function index()
    {
        $data['saham'] = $this->MSaham->getAll();
        loadPage('saham/index', $data);
    }

    public function tambah($id = null)
    {

        if ($id == null) {
            if (count($_POST)) {
                $this->form_validation->set_rules('saham', '', 'trim|required');
                if ($this->form_validation->run() == false) {
                    $errors = $this->form_validation->error_array();
                    $this->session->set_flashdata('errors', $errors);
                    redirect(current_url());
                } else {

                    $saham = $this->input->post('saham');
                    $nilai = $this->input->post('nilai');

                    $this->MSaham->saham = $saham;
                    if ($this->MSaham->insert() == true) {
                        $success = false;
                        $kdSaham = $this->MSaham->getLastID()->kdSaham;
                        foreach ($nilai as $item => $value) {
                            $this->MNilai->kdSaham= $kdSaham;
                            $this->MNilai->kdKriteria = $item;
                            $this->MNilai->nilai = $value;
                            if ($this->MNilai->insert()) {
                                $success = true;
                            }
                        }
                        if ($success == true) {
                            $this->session->set_flashdata('message', 'Berhasil menambah data :)');
                            redirect('saham');
                        } else {
                            echo 'gagal';
                        }
                    }
                }
                //-----
            }else{
                $data['dataView'] = $this->getDataInsert();
                loadPage('saham/tambah', $data);
            }
        }else{
            if(count($_POST)){
                $kdSaham = $this->uri->segment(3, 0);
                dump($kdSaham);
                if($kdSaham > 0){
                    $saham = $this->input->post('saham');
                    $nilai = $this->input->post('nilai');
                    $where = array('kdSaham' => $kdSaham);
                    $this->MSaham->saham = $saham;
                    dump($saham);
                    if($this->MSaham->update($where) == true){
                        $success = false;
                        foreach ($nilai as $item => $value) {
                            $this->MNilai->kdSaham = $kdSaham;
                            $this->MNilai->kdKriteria = $item;
                            $this->MNilai->nilai = $value;
                            if ($this->MNilai->update()) {
                                $success = true;
                            }
                        }
                        if ($success == true) {
                            $this->session->set_flashdata('message', 'Berhasil mengubah data :)');
                            redirect('saham');
                        } else {
                            echo 'gagal';
                        }
                    }
                }
            }
            $data['dataView'] = $this->getDataInsert();
            $data['nilaiSaham'] = $this->MNilai->getNilaiBySaham($id);
            loadPage('saham/tambah', $data);
        }

    }

    private function getDataInsert()
    {
        $dataView = array();
        $kriteria = $this->MKriteria->getAll();
        foreach ($kriteria as $item) {
            $this->MSubKriteria->kdKriteria = $item->kdKriteria;
            $dataView[$item->kdKriteria] = array(
                'nama' => $item->kriteria,
                'data' => $this->MSubKriteria->getById()
            );
        }

        return $dataView;
    }

    public function delete($id)
    {
        if($this->MNilai->delete($id) == true){
            if($this->MUniversitas->delete($id) == true){
                $this->session->set_flashdata('message','Berhasil menghapus data :)');
                echo json_encode(array("status" => 'true'));
            }
        }
    }
}