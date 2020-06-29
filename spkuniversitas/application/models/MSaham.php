<?php

class MSaham extends CI_Model{

    public $kdSaham;
    public $saham;

    public function __construct(){
        parent::__construct();
    }

    private function getTable(){
        return 'saham';
    }

    private function getData(){
        $data = array(
            'saham' => $this->saham
        );

        return $data;
    }

    public function getAll()
    {
        $saham = array();
        $query = $this->db->get($this->getTable());
        if($query->num_rows() > 0){
            foreach ($query->result() as $row) {
                $saham[] = $row;
            }
        }
        return $saham;
    }


    public function insert()
    {
        $this->db->insert($this->getTable(), $this->getData());
        return $this->db->insert_id();
    }

    public function update($where)
    {
        $status = $this->db->update($this->getTable(), $this->getData(), $where);
        return $status;

    }

    public function delete($id)
    {
        $this->db->where('kdSaham', $id);
        return $this->db->delete($this->getTable());
    }

    public function getLastID(){
        $this->db->select('kdSaham');
        $this->db->order_by('kdSaham', 'DESC');
        $this->db->limit(1);
        $query = $this->db->get($this->getTable());
        return $query->row();
    }


}